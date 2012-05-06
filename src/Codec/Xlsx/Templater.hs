{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Templater(
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
import           Data.Text hiding (map)
import           Data.Time.LocalTime
import           Text.Parsec
import           Text.Parsec.Text

data TemplateSettings = TemplateSettings { }
data TemplateValue = TplText Text | TplDouble Double | TplLocalTime LocalTime
type DataRow = M.Map Text TemplateValue

data Converter = Match Text | PassThrough
               deriving Show

data TmplCell = TmplCell{ converter :: Converter
                        , srcCell :: Cell }
              deriving Show

buildTemplate :: [Cell] -> [TmplCell]
buildTemplate = map build
  where
    build cell = TmplCell{ converter = conv cell
                         , srcCell = cell}
    conv cell =
      case cellValue cell of
        Just v@(CellText t)  -> either (const $ PassThrough) Match (getVar t)
        Nothing -> PassThrough
    getVar = parse varParser "not gonna show up"
    varParser :: Parser Text
    varParser = do
      string "{{"
      name <- many1 $ noneOf "}"
      string "}}"
      return $ pack name

applyTemplate :: [TmplCell] -> DataRow -> [Cell]
applyTemplate t r = map transform t
  where
    transform tc =
      case converter tc of
        Match k     -> (srcCell tc){cellValue = tpl2xlsx $ fromJust $ M.lookup k r}
        PassThrough -> srcCell tc
    tpl2xlsx (TplText t) = Just $ CellText t
    tpl2xlsx (TplDouble d) = Just $ CellDouble d
    tpl2xlsx (TplLocalTime t) = Just $ CellLocalTime t

run :: FilePath -> FilePath -> TemplateSettings -> [DataRow] -> IO ()
run tp op s d = do
  x@Xlsx{styles=Styles sbs} <- xlsx tp
  templateRow <- sheet x 0 ["A","B","C","D"] $$ CL.head
  
--  putStrLn $ "tmpl:" ++ show 

  let tpl = buildTemplate $ fromJust templateRow
      d' = map (applyTemplate tpl) d
      --map mapRow d
--      mapRow :: DataRow -> [Cell]
--      mapRow x = [maybe Nothing (Just CellText) $ M.lookup (pack "x") x]
  writeXlsxStyles op sbs d'
