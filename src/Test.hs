{-# LANGUAGE OverloadedStrings #-}
module Test where

import Codec.Xlsx.Templater
import Data.Map
import Data.Time.Calendar
import Data.Time.LocalTime

test = run "tmpl.xlsx" "tmpl-out.xlsx" common ((TemplateSettings Rows 3), d)
  where
    common = fromList [("common1", TplText "common first value"), ("commonDbl", TplDouble 42)]
    d =
      [fromList [("a", TplText "a value"),("b", TplText "b value"),("c", TplText "c value")],
       fromList [("a", TplText "a value'"),("b", TplDouble 1.23456),("c", TplLocalTime $ LocalTime (fromGregorian 2012 08 13) (TimeOfDay 12 13 14))]]
