;; # Excel Utilities
;; This file provides utility functions for reading `.xlsx` files.
;; It's a wrapper around a small part of the
;; [Apache POI project](http://poi.apache.org).
;; See the `incanter-excel` module from the
;; [Incanter](https://github.com/liebke/incanter) project for a more
;; thorough implementation.
;; TODO: Dates are not handled.

;; The function definitions progress from handling cells to rows, to sheets,
;; to workbooks.
(ns ontodev.excel
  (:require [clojure.tools.logging :as log]
            [clojure.string :as string]
            [clojure.java.io :as io]
            [java-time :as time])
  (:import
   (org.apache.poi.ss.usermodel Cell CellType Row Row$MissingCellPolicy
                                Sheet Workbook WorkbookFactory
                                DateUtil DataFormatter FormulaEvaluator)))

;; ## Cells
;; I've found it hard to trust the Cell Type and Cell Style for data such as
;; integers. In this version of the code I'm converting each cell to STRING
;; type before reading it as a string and returning the string value.
;; This should be the literal value typed into the cell, except in the case
;; of formulae where it should be the result.
;; Conversion of the strings to other data types should be done as an
;; additional step.
;;
;; With POI 4.1.2 Cell Type and formatting seem to work fine, including date types.

(defn version
  "Return version details, for now manually"
  []
  "ndevreeze/excel v0.3.2-SNAPSHOT, 2024-03-30 11:42, incl handling of / in to-keyword")

(defn get-cell-string-value
  "Get the value of a cell as a string, by changing the cell type to 'string'
   and then changing it back.
   optional cell-type is not used here."
  ([cell] (get-cell-string-value cell nil))
  ([cell cell-type]
   (let [ct    (.getCellType cell)
         cf    (if (= ct CellType/FORMULA) (.getCellFormula cell))
         _     (.setCellType cell CellType/STRING)
         value (.getStringCellValue cell)]
     (if (= ct CellType/FORMULA)
       (.setCellFormula cell cf)
       (.setCellType cell ct))
     value)))

;; ## Rows
;; Rows are made up of cells. We consider the first row to be a header, and
;; translate its values into keywords. Then we return each subsequent row
;; as a map from keys to cell values.

#_(defn to-keyword
    "Take a string and return a properly formatted keyword."
    [s]
    (-> (or s "")
        string/trim
        string/lower-case
        (string/replace #"\s+" "-")
        keyword))

(defn to-keyword
  "Take a string and return a properly formatted keyword.
   Replace all characters outside of letter, digits and underscore"
  [s]
  (-> (or s "")
      string/trim
      string/lower-case
      (string/replace #"\s+" "-")
      (string/replace #"[^A-Za-z0-9_]" "-")
      keyword))

(defn int-value?
  "Return true iff the actual value is an integer.
   Datatype could be a float/double"
  [v]
  (zero? (- v (int v))))

(defn cell-value-numeric
  "Return numeric cell-value, could be a date/time"
  [^Cell cell]
  (let [val (.getNumericCellValue cell)]
    (if (DateUtil/isCellDateFormatted cell)
      (cond (int-value? val) (time/local-date (.getLocalDateTimeCellValue cell))
            (< val 1)        (time/local-time (.getLocalDateTimeCellValue cell))
            :else            (.getLocalDateTimeCellValue cell))
      ;; else, not a date-time
      (if (int-value? val)
        (int (.getNumericCellValue cell))
        (.getNumericCellValue cell)))))

(defn cell-value
  "Return cell-value based on proper getter based on cell-value"
  ([^Cell cell] (cell-value cell (.getCellType cell)))
  ([^Cell cell ^CellType cell-type]
   (condp = cell-type
     CellType/BLANK nil
     CellType/STRING (.getStringCellValue cell)
     CellType/NUMERIC (cell-value-numeric cell)
     CellType/BOOLEAN (.getBooleanCellValue cell)
     CellType/FORMULA (cell-value cell (.getCachedFormulaResultType cell))
     CellType/ERROR (.getErrorCellValue cell)
     "unsupported")))

(defn cell-value-formatted
  "Return cell-value based on proper getter based on cell-value.
   Some Excel functions are not implemented (eg GETPIVOTDATA),
  fallback to cell-value"
  ([^FormulaEvaluator evaluator ^DataFormatter data-formatter ^Cell cell]
   (cell-value-formatted evaluator data-formatter cell (.getCellType cell)))
  ([^FormulaEvaluator evaluator ^DataFormatter data-formatter ^Cell cell ^CellType cell-type]
   (letfn [(cell-value-formula-formatted
             [^FormulaEvaluator evaluator ^DataFormatter data-formatter ^Cell cell]
             (try
               (cell-value-formatted evaluator data-formatter cell (.getCachedFormulaResultType cell))
               (catch Exception e
                 (cell-value cell (.getCachedFormulaResultType cell)))))]
     (condp = cell-type
       CellType/BLANK nil
       CellType/STRING (.getStringCellValue cell) ;; formatter needed here?
       CellType/NUMERIC (.formatCellValue data-formatter cell evaluator)
       CellType/BOOLEAN (.getBooleanCellValue cell)
       CellType/FORMULA (cell-value-formula-formatted evaluator data-formatter cell)
       CellType/ERROR (.getErrorCellValue cell)
       "unsupported"))))

;; Note: it would make sense to use the iterator for the row. However that
;; iterator just skips blank cells! So instead we use an uglier approach with
;; a list comprehension. This relies on the workbook's setMissingCellPolicy
;; in `load-workbook`.
;; See `incanter-excel` and [http://stackoverflow.com/questions/4929646/how-to-get-an-excel-blank-cell-value-in-apache-poi]()
;; This still holds for POI 4.1.2.
(defn read-row
  "Read all the cells in a row (including blanks) and return a list of values."
  ([^Row row] (read-row get-cell-string-value row))
  ([cell-value-fn ^Row row]
   (for [i (range 0 (.getLastCellNum row))]
     (cell-value-fn (.getCell row i)))))

(defn cell-value-fn
  "Return function to determine Excel cell values based on given options.
   opt - cmdline options.
   workbook - the Excel workbook object, used to create the function.
   if `--values values` is given, use the underlying values.
   if `--values formatted` is given, use the cells as formatted (visbly) in the excel sheet.
   The returned function should accept 1 or 2 arguments: the first is the cell, the optional
   second is the cell-type."
  [opt ^Workbook workbook]
  (case (:values opt)
    :values cell-value
    :formatted (partial cell-value-formatted
                        (.. workbook getCreationHelper createFormulaEvaluator)
                        (DataFormatter.))
    :strings get-cell-string-value
    ;; default as string values, like original.
    get-cell-string-value))

(defn read-sheet
  "Given a workbook with an optional sheet name (default is 'Sheet1') and
   and optional header row number (default is '1'),
   return the data in the sheet as a vector of maps
   using the headers from the header row as the keys.
   Use :values key in opt (map) to determine how to determine the values:
     :strings   - the default, all values are returned as strings
     :values    - the actual values with correct datatype, including dates/times based on cell-formatting
     :formatted - the formatted values"
  ([^Workbook workbook] (read-sheet {} workbook "Sheet1" 1))
  ([opt ^Workbook workbook] (read-sheet opt workbook "Sheet1" 1))
  ([opt ^Workbook workbook ^String sheet-name] (read-sheet opt workbook sheet-name 1))
  ([opt ^Workbook workbook ^String sheet-name n-header-rows]
   (let [sheet          (.getSheet workbook sheet-name)
         rows           (->> sheet (.iterator) iterator-seq (drop (dec n-header-rows)))
         cell-fn        (cell-value-fn opt workbook)
         read-row-fn    (partial read-row cell-fn)
         read-header-fn (partial read-row get-cell-string-value)
         headers        (map to-keyword (read-header-fn (first rows)))
         data           (map read-row-fn (rest rows))]
     (vec (map (partial zipmap headers) data)))))

(defn list-sheets
  "Return a list of all sheet names."
  [workbook]
  (for [i (range (.getNumberOfSheets workbook))]
    (.getSheetName workbook i)))

(defn sheet-headers
  "Returns the headers (in their original forms, not as keywords) for a given sheet."
  [workbook sheet-name]
  (let [sheet (.getSheet workbook sheet-name)
        rows (->> sheet (.iterator) iterator-seq)]
    (read-row (first rows))))

;; ## Workbooks
;; An `.xlsx` file contains one workbook with one or more sheets.

(defn load-workbook
  "Load a workbook from a string path."
  [path]
  (log/debugf "Loading workbook:" path)
  (doto (WorkbookFactory/create (io/input-stream path))
    (.setMissingCellPolicy Row$MissingCellPolicy/CREATE_NULL_AS_BLANK)))
