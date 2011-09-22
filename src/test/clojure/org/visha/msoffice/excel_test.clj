;; -*- mode: clojure; coding: utf-8 -*-
(ns org.visha.msoffice.excel-test
  (:use [clojure.test]
        [org.visha.msoffice.excel])
  (:import (org.apache.poi.ss.usermodel Workbook)))

(deftest test_workbook-seq []
  (let [w (workbook (ClassLoader/getSystemResourceAsStream "sample.xls"))]
    (is (isa? (class w) Workbook))
  ))

(run-tests)
