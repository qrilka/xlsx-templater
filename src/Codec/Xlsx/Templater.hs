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
import           Data.List
import qualified Data.Map as M
import           Data.Maybe
import           Data.Text (Text, pack)
import           Data.Time.LocalTime
import           Text.Parsec
import           Text.Parsec.Text


data Orientation =  Rows | Columns
                 deriving (Show, Eq)
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

replacePlaceholders :: [[Maybe CellData]] -> TemplateDataRow -> [[Maybe CellData]]
replacePlaceholders d tdr = map (map $ fmap replace) d
  where
    replace :: CellData -> CellData
    replace cd@CellData{cdValue=Just (CellText t)} =
      either (const cd) (\ph -> cd{cdValue=Just (phValue ph)}) (getVar t)
    replace cd = cd
    phValue ph = maybe (CellText ph) tpl2xlsx (M.lookup ph tdr)

getVar = parse varParser "unnecessary error"
  where
    varParser = do
      string "{{"
      name <- many1 $ noneOf "}"
      string "}}"
      return $ pack name

buildTemplate :: Int -> [Maybe CellData] -> [TplCell]
buildTemplate x = map build
  where
    build cd = TplCell{ tplConverter = conv cd
                      , tplSrc       = cd
                      , tplX         = x}
    conv (Just CellData{cdValue=Just (CellText t)}) = either (const PassThrough) Match (getVar t)
    conv _ = PassThrough

applyTemplate :: [TplCell] -> TemplateDataRow -> [Maybe CellData]
applyTemplate t r = map transform t
  where
    transform tc = case tplConverter tc of
      Match k     -> do
        cd <- tplSrc tc
        v <- M.lookup k r
        return cd{cdValue = Just (tpl2xlsx v)}
      PassThrough -> tplSrc tc

fixColumns :: [ColumnsWidth] -> Int -> Int -> [ColumnsWidth]
fixColumns cw c n = prolog ++ dataepilog
  where
    (prolog, rest) = span ((<c) . cwMax) cw
    dataepilog = case rest of
      [] -> []
      (dCW : rest') -> fixD dCW : fixEpilog rest'
    fixD (ColumnsWidth dMin dMax width) = ColumnsWidth dMin (dMax + n - 1) width
    fixEpilog = map (\(ColumnsWidth dMin dMax width) -> ColumnsWidth (dMin + n - 1) (dMax + n - 1) width)

fixRowHeights :: RowHeights -> Int -> Int -> RowHeights
fixRowHeights rh r n = insertCopies $ shift removeOriginal
  where
    original = M.lookup r rh
    removeOriginal = M.delete r rh
    shift = M.mapKeys (\x -> if x > r then x + n - 1 else x)
    insertCopies m = case original of
      Just h -> foldr (\x m -> M.insert x h m) m [r..(r + n -1)]
      Nothing -> m


runSheet :: Xlsx -> Int -> (TemplateDataRow, TemplateSettings, [TemplateDataRow]) -> IO Worksheet
runSheet x n (cdr, ts, d) = do
  ws <- sheetMap x n
  let
    templateRows = if tsOrientation ts == Columns then transpose $ toList ws else toList ws
    repeatRow = tsRepeated ts
    (prolog, templateRow : epilog) = splitAt repeatRow templateRows
    tpl = buildTemplate repeatRow templateRow
    prolog' = replacePlaceholders prolog cdr
    n = length d
    d' = map (applyTemplate tpl) d
    epilog' = replacePlaceholders epilog cdr
    output = concat [prolog', d', epilog']
    result = if tsOrientation ts == Columns then transpose output else output
    (cw, rh) = if tsOrientation ts == Columns 
                 then (fixColumns (wsColumns ws) (repeatRow + 1) n, wsRowHeights ws) 
                 else (wsColumns ws, fixRowHeights (wsRowHeights ws) (repeatRow + 1) n)
    in
   return $ fromList (wsName ws) cw rh result


run :: FilePath -> FilePath -> [(TemplateDataRow, TemplateSettings, [TemplateDataRow])] -> IO ()
run tp op options = do
  x@Xlsx{styles=Styles sbs} <- xlsx tp
  out <- mapM (uncurry (runSheet x)) $ zip [0..] options
  writeXlsxStyles op sbs out
