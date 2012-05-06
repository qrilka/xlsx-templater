{-# LANGUAGE OverloadedStrings #-}
module Test where

import Data.Map
import Data.Xlsx.Templater

test = run "tmpl.xlsx" "tmpl-out.xlsx" TemplateSettings d
  where
    d = [fromList [("a", "a value"),
                   ("b", "b value"),
                   ("c", "c value")]]
