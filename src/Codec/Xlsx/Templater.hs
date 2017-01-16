{-# LANGUAGE OverloadedStrings #-}
module Codec.Xlsx.Templater(
  Orientation(..),
  TemplateSettings(..),
  TemplateDataRow,
  applyTemplateOnCell,
  applyTemplateOnSheet,
  applyTemplateOnXlsx,
  run
  ) where

import           Codec.Xlsx.Types hiding (Orientation)
import           Codec.Xlsx.Parser hiding (ParseError)
import           Codec.Xlsx.Writer
import           Control.Lens hiding (noneOf, op)
import qualified Data.ByteString.Lazy as L
import qualified Data.Map as M
import           Data.Text (Text, pack)
import           Data.Time.Clock.POSIX
import           Text.Parsec
import           Text.Parsec.Text()

data Orientation =  Rows | Columns deriving (Show, Eq)

data TemplateSettings = TemplateSettings { tsOrientation :: Orientation
                                         , tsRepeated    :: Int         -- ^ repeated row/column (depending on 'tsOrientation')
                                                                        -- starting at 0.
                                         }


-- | data row as a map from template variable name to a 'CellValue'
type TemplateDataRow = M.Map Text CellValue

applyTemplateOnCell :: TemplateDataRow -> Cell -> Cell
applyTemplateOnCell tdr cd@Cell{_cellValue=Just (CellText t)} =
    either (const cd) (\ph -> cd{_cellValue=Just (phValue ph)}) (getVar t)
  where
    phValue ph = maybe (CellText ph) id (M.lookup ph tdr)
applyTemplateOnCell _   cd = cd

getVar :: Text -> Either ParseError Text
getVar = parse varParser "unnecessary error"
  where
    varParser = string "{{" *> (pack <$> many1 (noneOf "}")) <* string "}}"

fixColumns :: [ColumnsWidth] -> Int -> Int -> [ColumnsWidth]
fixColumns cw c n = prolog ++ dataepilog
  where
    (prolog, rest) = span ((<c) . cwMax) cw
    dataepilog = case rest of
      [] -> []
      (dCW : rest') -> fixD dCW : map fixEpilog rest'
    fixD (ColumnsWidth dMin dMax width style) = ColumnsWidth dMin (dMax + n - 1) width style
    fixEpilog (ColumnsWidth dMin dMax width style) = ColumnsWidth (dMin + n - 1) (dMax + n - 1) width style
{- ColumnsWidth is not using lenses :(
    fixD = cwMax +~ (n - 1)
    fixEpilog = cwMin +~ (n - 1)
              . cwMax +~ (n - 1)
-}

fixRowHeights :: M.Map Int RowProperties -> Int -> Int -> M.Map Int RowProperties
fixRowHeights rh r n = insertCopies $ shift removeOriginal
  where
    original = M.lookup r rh
    removeOriginal = M.delete r rh
    shift = M.mapKeys (\x -> if x > r then x + n - 1 else x)
    insertCopies m = case original of
      Just h -> foldr (\x m' -> M.insert x h m') m [r..(r + n -1)]
      Nothing -> m

applyTemplateOnSheet :: (TemplateDataRow, TemplateSettings, [TemplateDataRow]) -> Worksheet -> Worksheet
applyTemplateOnSheet (cdr, ts, d) ws =
    ws & wsColumns           .~ cw
       & wsRowPropertiesMap  .~ rh
       & wsCells             %~ applyTemplateOnCellMap
  where
    mayTranspose =
      case tsOrientation ts of
        Columns -> each . _1 %~ (\(x,y) -> (y,x))
        _       -> id
    repeatRow = tsRepeated ts + 1
    n = length d
    applyTemplateOnCellMap =
      M.fromList . mayTranspose . concatMap applyTemplateOnCellPos . mayTranspose . M.toList
    applyTemplateOnCellPos ((row, col), cell)
      | row < repeatRow = [ ((row,     col), applyTemplateOnCell cdr cell) ]
      | row > repeatRow = [ ((row + n, col), applyTemplateOnCell cdr cell) ]
      | otherwise       = [ ((row + i, col), applyTemplateOnCell dc  cell)
                          | (i, dc) <- zip [0 ..] d ]
    (cw, rh) = if tsOrientation ts == Columns
                 then (fixColumns (ws ^. wsColumns) repeatRow n, ws ^. wsRowPropertiesMap)
                 else (ws ^. wsColumns, fixRowHeights (ws ^. wsRowPropertiesMap) repeatRow n)

applyTemplateOnXlsx :: [(TemplateDataRow, TemplateSettings, [TemplateDataRow])] -> Xlsx -> Xlsx
applyTemplateOnXlsx options x =
  x & xlSheets .~ zipWith (\(nm,sh) opt -> (nm, applyTemplateOnSheet opt sh)) (x ^. xlSheets) options

-- | template runner: reads template, constructs new xlsx file based on template data and template settings
run :: FilePath -> FilePath -> [(TemplateDataRow, TemplateSettings, [TemplateDataRow])] -> IO ()
run tp op options = do
  ct <- getPOSIXTime
  L.writeFile op . fromXlsx ct . applyTemplateOnXlsx options . toXlsx =<< L.readFile tp
