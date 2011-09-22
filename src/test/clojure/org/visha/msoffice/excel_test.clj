;; -*- mode: clojure; coding: utf-8 -*-
(ns org.visha.msoffice.excel-test
  (:use [clojure.test]
        [org.visha.msoffice.excel])
  (:import (org.apache.poi.ss.usermodel Workbook)))

(deftest test_xls-workbook-seq []
  (let [w (workbook (ClassLoader/getSystemResourceAsStream "sample.xls"))]
    (is (isa? (class w) Workbook))
    (is (= (class w) org.apache.poi.hssf.usermodel.HSSFWorkbook))
    ))

(deftest test_xlsx-workbook-seq []
  (let [w (workbook (ClassLoader/getSystemResourceAsStream "sample.xlsx"))]
    (is (isa? (class w) Workbook))
    (is (= (class w) org.apache.poi.xssf.usermodel.XSSFWorkbook))
    ))

(run-tests)
