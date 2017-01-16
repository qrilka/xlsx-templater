{-# LANGUAGE OverloadedStrings #-}
import Codec.Xlsx.Types
import Codec.Xlsx.Templater
import Data.Map
import Data.Fixed (Pico)
import Data.String
import Data.Text (Text)
import Data.Time.Calendar
import Data.Time.LocalTime

{-
test :: IO ()
test = run "tmpl.xlsx" "tmpl-out.xlsx" [(empty, TemplateSettings Rows 0, d1),
                                        (common, TemplateSettings Rows 3, d2),
                                        (empty, TemplateSettings Columns 1, d3)]
  where
    d1 = [fromList [("a", txt "first row"), ("b", txt "b1")],
          fromList [("a", txt "second row should be here!"), ("b", txt "b2")]]
    common = fromList [("common1", txt "common first value"), ("commonDbl", n 42)]
    d2 =
      [fromList [("a", txt "a value"),("b", txt "b value"),("c", txt "c value")],
       fromList [("a", txt "a value'"),("b", n 1.23456),("c", dt 13 08 2012 12 13 14)]]
    d3 = [fromList [("x", txt "first column"), ("z", txt "z1")],
          fromList [("x", txt "second column should be here!"), ("z", txt "z2")],
          fromList [("x", txt "column #3"), ("z", txt "z3")]]
-}

n :: Double -> CellValue
n = CellDouble

txt :: Text -> CellValue
txt = CellText

dt :: Integer -> Int -> Int -> Int -> Int -> Pico -> CellValue
dt d' m y h mi s = txt . fromString . show $ LocalTime (fromGregorian d' m y) (TimeOfDay h mi s)

d :: Integer -> Int -> Int -> CellValue
d d' m y = dt d' m y 0 0 0

t :: Int -> Int -> Pico -> CellValue
t = dt 2000 1 1

repairData :: [Map Text CellValue]
repairData = replicate 10000 $ fromList
             [("N", n 23456), ("Org", txt "XYZ Co."), ("Date", d 2012 1 13), ("Time", t 10 00 00),
              ("CarNum", txt "X666YZ777"), ("SaleDate", d 2009 12 07), ("VIN", txt "WWWZZZ6KZBW666666"),
              ("Run", n 77777), ("Brand", txt "BELAZ"), ("Model", txt "Tachila"),
              ("Problem", txt "Something broken"), ("ProblemSystem", txt "Electornics"),
              ("ProblemDetail", txt "No power"), ("ProblemCause", txt "Bug found"),
              ("Weather", txt "Windy with some clouds"), ("HelpResult", txt "Removed bug"),
              ("Status", txt "Fixed"), ("AppealType", txt "Evacuation"), ("Evacuator", txt "BORYA Ltd."),
              ("EvacuationCity", txt "Chicago"), ("SaleDealer", txt "Some greate Car Dealer"),
              ("ProblemAddress", txt "Somewhere over the rainbow"), ("CustomerFirstName", txt "John"),
              ("CustomerSurname", txt "Dow"), ("OrderNumber", txt "some"), ("RepairFinishDate",  d 2012 2 1),
              ("CountryEvacuatorRun", txt "-"), ("CustomerPhone", txt "+7 (495) 666 6666"),
              ("Comment", txt "some notes here"), ("Cost", n 6666.66), ("CostExplained", txt "10 * 666,666")]

main :: IO ()
main = {-test >>-} run "data/Repairs-template.xlsx" "data/Repairs.xlsx" [(empty, TemplateSettings Rows 1, repairData)]
