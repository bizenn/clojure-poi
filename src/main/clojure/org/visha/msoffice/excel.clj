;; -*- mode: clojure; coding: utf-8 -*-
(ns org.visha.msoffice.excel
  "Wrapper library on Apache POI HSSF and XSSF."
  (:use [clojure.java.io :only [file input-stream]])
  (:import (org.apache.poi.ss.usermodel WorkbookFactory Cell)
           (java.io File InputStream)))

(defmulti workbook
  "Read from String(filename)/File/InputStream and return a workbook."
  class)
(defmethod workbook InputStream [input]
  (with-open [in input]
    (WorkbookFactory/create in)))
(defmethod workbook File [file]
  (workbook (input-stream file)))
(defmethod workbook String [filename]
  (workbook (file filename)))

(defn workbook-seq
  "Handle a workbook as lazy sequence of sheets"
  ([wb] (workbook-seq wb (.getNumberOfSheets wb) 0))
  ([wb number-of-sheets index]
     (if (>= index number-of-sheets)
       ()
       (lazy-seq
        (cons (.getSheetAt wb index) (workbook-seq wb number-of-sheets (inc index)))))))

(defn sheet-seq
  "Handle a sheet as lazy sequence of rows."
  [s] (iterator-seq (.iterator s)))

(defn row-seq
  "Handle a row as lazy sequence of cells."
  [r] (iterator-seq (.iterator r)))

(defmulti cell-value
  "Get cell value as a proper type one."
  (fn [c] (if (nil? c) nil (.getCellType c))))
(defmethod cell-value nil [cell]
  nil)
(defmethod cell-value Cell/CELL_TYPE_NUMERIC [cell]
  (.getNumericCellValue cell))
(defmethod cell-value Cell/CELL_TYPE_STRING [cell]
  (.getStringCellValue cell))
(defmethod cell-value Cell/CELL_TYPE_FORMULA [cell]
  (.getFormulaCellValue cell))
(defmethod cell-value Cell/CELL_TYPE_BLANK [cell]
  nil)
(defmethod cell-value Cell/CELL_TYPE_BOOLEAN [cell]
  (.getBooleanCellValue cell))
(defmethod cell-value Cell/CELL_TYPE_ERROR [cell]
  (.getErrorCellValue cell))
