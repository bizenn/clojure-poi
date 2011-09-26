;; -*- mode: clojure; coding: utf-8 -*-
(ns org.visha.msoffice.excel-test
  (:use [clojure.test]
        [org.visha.msoffice.excel])
  (:import (org.apache.poi.ss.usermodel Workbook Sheet Row Cell)))

(def sample-values
  '((1.0 "テレキャスター" "Fender")
    (2.0 "ストラトキャスター" "Fender")
    (3.0 "レスポール・スタンダード" "Gibson")))

(deftest test_cell-value []
  (is (= (nil? (cell-value nil)))))

(deftest test_xls-workbook-seq []
  (let [w (workbook (ClassLoader/getSystemResourceAsStream "sample.xls"))]
    (is (isa? (class w) Workbook))
    (is (= (class w) org.apache.poi.hssf.usermodel.HSSFWorkbook))
    (let [sheets (workbook-seq w)]
      (is (seq? sheets))
      (doseq [sheet sheets]
        (is (isa? (class sheet) Sheet))
        (is (= (class sheet) org.apache.poi.hssf.usermodel.HSSFSheet))
        (let [rows (sheet-seq sheet)]
          (is (seq? rows))
          (doseq [[row row-values] (map list rows sample-values)]
            (is (isa? (class row) Row))
            (is (= (class row) org.apache.poi.hssf.usermodel.HSSFRow))
            (let [cells (row-seq row)]
              (is (seq? cells))
              (doseq [[cell value] (map list cells row-values)]
                (is (isa? (class cell) Cell))
                (is (= (class cell) org.apache.poi.hssf.usermodel.HSSFCell))
                (is (= (cell-value cell) value))
                ))))))))

(deftest test_xlsx-workbook-seq []
  (let [w (workbook (ClassLoader/getSystemResourceAsStream "sample.xlsx"))]
    (is (isa? (class w) Workbook))
    (is (= (class w) org.apache.poi.xssf.usermodel.XSSFWorkbook))
    (let [sheets (workbook-seq w)]
      (is (seq? sheets))
      (doseq [sheet sheets]
        (is (isa? (class sheet) Sheet))
        (is (= (class sheet) org.apache.poi.xssf.usermodel.XSSFSheet))
        (let [rows (sheet-seq sheet)]
          (is (seq? rows))
          (doseq [[row row-values] (map list rows sample-values)]
            (is (isa? (class row) Row))
            (is (= (class row) org.apache.poi.xssf.usermodel.XSSFRow))
            (let [cells (row-seq row)]
              (is (seq? cells))
              (doseq [[cell value] (map list cells row-values)]
                (is (isa? (class cell) Cell))
                (is (= (class cell) org.apache.poi.xssf.usermodel.XSSFCell))
                (is (= (cell-value cell) value))
                ))))))))

(run-tests)
