{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Templater(
  Orientation(..),
  TemplateSettings(..),
  TemplateValue(..),
  run
  ) where

import           Codec.Xlsx
import           Codec.Xlsx.Parser
import           Codec.Xlsx.Writer
import           Control.Applicative ((<$>))
import           Data.Conduit
import qualified Data.Conduit.List as CL
import           Data.Default (Default (..))
import qualified Data.Map as M
import           Data.Maybe
import           Data.Text (Text, pack)
import           Data.Time.LocalTime
import           Text.Parsec
import           Text.Parsec.Text

data Orientation =  Rows | Columns
data TemplateSettings = TemplateSettings { tsOrientation :: Orientation
                                         , tsRepeated    :: Int
                                         }
data TemplateValue = TplText Text | TplDouble Double | TplLocalTime LocalTime
                   deriving Show
type TemplateDataRow = M.Map Text TemplateValue

data Converter = Match Text | PassThrough
               deriving Show

data TplCell = TplCell{ tplConverter :: Converter
                      , tplSrc       :: Maybe CellData
                      , tplX         :: Int
                      }
              deriving Show

tpl2xlsx :: TemplateValue -> CellValue
tpl2xlsx (TplText t) = CellText t
tpl2xlsx (TplDouble d) = CellDouble d
tpl2xlsx (TplLocalTime t) = CellLocalTime t

type TplDataCell = (Int, Int, Maybe CellData)

replacePlaceholders :: Int -> [[TplDataCell]] -> TemplateDataRow -> (Int, [[TplDataCell]])
replacePlaceholders n d tdr = countMap n (\n -> map $ replace n) d
  where
    replace :: Int -> TplDataCell -> TplDataCell
    replace y (x, _, src@(Just cd@CellData{cdValue=CellText t})) = 
      either (const (x, y, src)) (\ph -> (x, y, Just cd{cdValue=phValue ph})) (getVar t)
    replace y (x, _, src) = (x, y, src)
    phValue ph = maybe (CellText ph) tpl2xlsx (M.lookup ph tdr)

countMap :: Int -> (Int -> a -> b) -> [a] -> (Int, [b])
countMap n f [] = (n, [])
countMap n f l = (n' + 1, transformed)
  where
    n' = fst $ last numbered
    numbered = zip [n..] l
    transformed = map (uncurry f) numbered -- (\(n, a) -> f n a)

getVar = parse varParser "unnecessary error"
  where
    varParser = do
      string "{{"
      name <- many1 $ noneOf "}"
      string "}}"
      return $ pack name

buildTemplate :: [TplDataCell] -> [TplCell]
buildTemplate = map build
  where
    build (x, _, cd) = TplCell{ tplConverter = conv cd
                              , tplSrc       = cd
                              , tplX         = x}
    conv cd =
      case cd of
        Just CellData{cdValue=CellText t}  -> either (const $ PassThrough) Match (getVar t)
        Nothing -> PassThrough

applyTemplate :: [TplCell] -> Int -> TemplateDataRow -> [TplDataCell]
applyTemplate t y r = map transform t
  where
    transform tc = (tplX tc, y,
                    case tplConverter tc of
                      Match k     -> fmap (\cd -> cd{cdValue = tpl2xlsx $ fromJust $ M.lookup k r}) (tplSrc tc)
                      PassThrough -> tplSrc tc)


map2matrix :: (Maybe ((Int,Int), (Int,Int), M.Map (Int,Int) CellData)) -> [[TplDataCell]]
map2matrix Nothing = error "invalid template"
map2matrix (Just ((xmin,xmax), (ymin,ymax), m)) = [[lookupCell x y | x <- [1..xmax]] | y <- [1..ymax]]
  where
    lookupCell x y = (x, y, M.lookup (x, y) m)

tplData2Cell :: TplDataCell -> Cell
tplData2Cell (x, y, cd) = 
  case cd of
    Nothing -> 
      Cell{cellIx=(col,row), cellStyle=Nothing, cellValue=Nothing}
    Just CellData{cdValue=v,cdStyle=s} -> 
      Cell{cellIx=(col,row), cellStyle=s, cellValue=Just v}
  where
    col = int2col x
    row = y

runSheet x n (cdr, ts, d) = do
  templateRows <- map2matrix <$> sheetMap x n
  let
    (prolog, templateRow : epilog) = splitAt (tsRepeated ts) templateRows -- :: [[Cell]])
    tpl = buildTemplate templateRow 
    (n,  prolog') = replacePlaceholders 1 prolog cdr 
    (n', d') = countMap n (applyTemplate tpl) d 
    (_,  epilog') = replacePlaceholders n' epilog cdr in
    return $ map (map tplData2Cell) $ concat [prolog', d', epilog']


run :: FilePath -> FilePath -> [(TemplateDataRow, TemplateSettings, [TemplateDataRow])] -> IO ()
run tp op options = do
  x@Xlsx{styles=Styles sbs} <- xlsx tp
  out <- mapM (\(n, opts) -> runSheet x n opts ) $ zip [0..] options
  writeXlsxStyles op sbs out
