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
type TemplateDataRow = M.Map Text TemplateValue

data Converter = Match Text | PassThrough
               deriving Show

data TplCell = TplCell{ converter :: Converter
                      , srcCell   :: Cell
                      }
              deriving Show

tpl2xlsx :: TemplateValue -> Maybe CellValue
tpl2xlsx (TplText t) = Just $ CellText t
tpl2xlsx (TplDouble d) = Just $ CellDouble d
tpl2xlsx (TplLocalTime t) = Just $ CellLocalTime t

replacePlaceholders :: [[Cell]] -> TemplateDataRow -> [[Cell]]
replacePlaceholders d tdr = map (map replace) d
  where
    replace :: Cell -> Cell
    replace c@Cell{cellValue=Just (CellText t)} = either (const c) (\ph -> c{cellValue=phValue ph}) (getVar t)
    replace c = c
    phValue ph = maybe Nothing tpl2xlsx (M.lookup ph tdr)

getVar = parse varParser "unnecessary error"
  where
    varParser = do
      string "{{"
      name <- many1 $ noneOf "}"
      string "}}"
      return $ pack name

buildTemplate :: [Cell] -> [TplCell]
buildTemplate = map build
  where
    build cell = TplCell{ converter = conv cell
                         , srcCell = cell}
    conv cell =
      case cellValue cell of
        Just v@(CellText t)  -> either (const $ PassThrough) Match (getVar t)
        Nothing -> PassThrough

applyTemplate :: [TplCell] -> TemplateDataRow -> [Cell]
applyTemplate t r = map transform t
  where
    transform tc =
      case converter tc of
        Match k     -> (srcCell tc){cellValue = tpl2xlsx $ fromJust $ M.lookup k r}
        PassThrough -> srcCell tc

runSheet x n (cdr, ts, d) = do
  templateRows <- sheet x n ["A","B","C","D"] $$ CL.consume
  let
    (prolog, templateRow : epilog) = splitAt (tsRepeated ts) (templateRows :: [[Cell]])
    tpl = buildTemplate templateRow in
    return (replacePlaceholders prolog cdr ++ map (applyTemplate tpl) d ++ replacePlaceholders epilog cdr)


run :: FilePath -> FilePath -> [(TemplateDataRow, TemplateSettings, [TemplateDataRow])] -> IO ()
run tp op options = do
  x@Xlsx{styles=Styles sbs} <- xlsx tp
  out <- mapM (\(n, opts) -> runSheet x n opts ) $ zip [0..] options
  writeXlsxStyles op sbs out
