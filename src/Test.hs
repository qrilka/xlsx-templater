{-# LANGUAGE OverloadedStrings #-}
module Test where

import Codec.Xlsx.Templater
import Data.Map
import Data.Time.Calendar
import Data.Time.LocalTime

test = run "tmpl.xlsx" "tmpl-out.xlsx" [(empty, TemplateSettings Rows 0, d1),
                                        (common, TemplateSettings Rows 3, d2)]
  where
    d1 = [fromList [("a", TplText "first row"), ("b", TplText "b1")],
          fromList [("a", TplText "second row should be here!"), ("b", TplText "b2")]]
    common = fromList [("common1", TplText "common first value"), ("commonDbl", TplDouble 42)]
    d2 =
      [fromList [("a", TplText "a value"),("b", TplText "b value"),("c", TplText "c value")],
       fromList [("a", TplText "a value'"),("b", TplDouble 1.23456),("c", TplLocalTime $ LocalTime (fromGregorian 2012 08 13) (TimeOfDay 12 13 14))]]
