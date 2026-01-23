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
            [flatland.ordered.map :as omap]
            ;; [flatland.ordered.set :as oset]
            [java-time :as time])
  (:import
   (org.apache.poi.ss.usermodel Cell CellType Row Row$MissingCellPolicy
                                Workbook WorkbookFactory
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

;; [2024-08-10 12:51] Issues with 5.3.0, create a test for this, hopefully fix.
(defn version
  "Return version details, for now manually"
  []
  "ndevreeze/excel v0.3.2-SNAPSHOT, [2025-10-30 19:56] with ordered-map")

#_(defn version
    "Return version details, for now manually"
    []
    "ndevreeze/excel v0.3.2-SNAPSHOT, [2024-10-30 18:25] with sorted-map")

(defn get-cell-string-value
  "Get the value of a cell as a string, by changing the cell type to 'string'
   and then changing it back.
   optional cell-type is not used here."
  ([cell] (get-cell-string-value cell nil))
  ([cell _cell-type]
   (let [ct    (.getCellType cell)
         cf    (when (= ct CellType/FORMULA) (.getCellFormula cell))
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
               (catch Exception _e
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

(defn combine-headers
  [& strings]
  (string/join " " strings))

(defn combine-row-headers
  [rows-header]
  (let [header-rows (map read-row rows-header)
        headers (apply map combine-headers header-rows)]
    (map to-keyword headers)))

#_(defn sorted-zipmap
    "zipmap, but put keys in alphabetical order"
    [keys vals]
    (into (sorted-map) (zipmap keys vals)))

;; this one is wrong, zipmap already changes the order, ordered-map is too late.
#_(defn ordered-zipmap
    "zipmap, but put keys in original order"
    [keys vals]
    (into (omap/ordered-map) (zipmap keys vals)))

#_(defn regular-zipmap
    "zipmap, but put keys in original order"
    [keys vals]
    (into {} (zipmap keys vals)))

(comment
  ;; create ordered-map where key-order is same as in keys.
  (def keys [:a :b :c :d :e])
  (def vals [1 2 3 4 5])
  #_(sorted-zipmap keys vals) ;; this is ok, but sorted, so another test:

  (def keys [:z :b :x :d :e])
  (def vals [1 2 3 4 5])
  #_(sorted-zipmap keys vals) ;; b, d, e, x, z, so indeed sorted. But want to keep orig order.

  (ordered-zipmap keys vals) ;; z, b, x, d, e, so ok

  ;; regular map, should mix things up:
  #_(regular-zipmap keys vals) ;; z, b, x, d, e, so also ok

  )

(defn ordered-zipmap
  "Returns an ordered-map with the keys mapped to the corresponding vals.
  zipmap, but put keys in original order"
  {:added "1.0"
   :static true}
  [keys vals]
  (loop [map (transient (omap/ordered-map))
         ks (seq keys)
         vs (seq vals)]
    (if (and ks vs)
      (recur (assoc! map (first ks) (first vs))
             (next ks)
             (next vs))
      (persistent! map))))

(comment
  (def m1 (ordered-zipmap [:a :b :c :d :e] [11 12 13 14 15]))
  (merge m1 {:rownr 22})
  (assoc m1 :rownr 22)
  )


;; 2026-01-23: Add n-drop-rows: do not use the first n rows at all.
(defn read-sheet
  "Given a workbook with an optional sheet name (default is 'Sheet1') and
   and optional header row number (default is '1'),
   return the data in the sheet as a vector of maps
   using the headers from the header row as the keys.
   Use :values key in opt (map) to determine how to determine the values:
     :strings   - the default, all values are returned as strings
     :values    - the actual values with correct datatype, including dates/times based on cell-formatting
     :formatted - the formatted values
   If n-header-rows > 1, use all of the headers to determine field-name"
  ([^Workbook workbook] (read-sheet {} workbook "Sheet1" 1))
  ([opt ^Workbook workbook] (read-sheet opt workbook "Sheet1" 1))
  ([opt ^Workbook workbook ^String sheet-name] (read-sheet opt workbook sheet-name 1))
  ([opt ^Workbook workbook ^String sheet-name n-header-rows] (read-sheet opt workbook sheet-name n-header-rows 0))
  ([opt ^Workbook workbook ^String sheet-name n-header-rows n-drop-rows]
   (let [sheet          (.getSheet workbook sheet-name)
         rows-all       (->> sheet (.iterator) iterator-seq)
         rows-all2      (drop n-drop-rows rows-all)
         rows-header    (take n-header-rows rows-all2)
         rows-data      (drop n-header-rows rows-all2)
         cell-fn        (cell-value-fn opt workbook)
         read-row-fn    (partial read-row cell-fn)
         headers        (combine-row-headers rows-header)
         data           (map read-row-fn rows-data)]
     (vec (map (partial ordered-zipmap headers) data)))))

#_(defn read-sheet
    "Given a workbook with an optional sheet name (default is 'Sheet1') and
   and optional header row number (default is '1'),
   return the data in the sheet as a vector of maps
   using the headers from the header row as the keys.
   Use :values key in opt (map) to determine how to determine the values:
     :strings   - the default, all values are returned as strings
     :values    - the actual values with correct datatype, including dates/times based on cell-formatting
     :formatted - the formatted values
   If n-header-rows > 1, use all of the headers to determine field-name"
    ([^Workbook workbook] (read-sheet {} workbook "Sheet1" 1))
    ([opt ^Workbook workbook] (read-sheet opt workbook "Sheet1" 1))
    ([opt ^Workbook workbook ^String sheet-name] (read-sheet opt workbook sheet-name 1))
    ([opt ^Workbook workbook ^String sheet-name n-header-rows]
     (let [sheet          (.getSheet workbook sheet-name)
           rows-all       (->> sheet (.iterator) iterator-seq)
           rows-header    (take n-header-rows rows-all)
           rows-data      (drop n-header-rows rows-all)
           cell-fn        (cell-value-fn opt workbook)
           read-row-fn    (partial read-row cell-fn)
           headers        (combine-row-headers rows-header)
           data           (map read-row-fn rows-data)]
       (vec (map (partial ordered-zipmap headers) data)))))

(comment
  ;; 2025-11-04: test with incidents, order of columns. Even with ordered-zipmap, it's still the wrong order.
  (do
    (def excel-path "/Users/nicodevreeze/projects/incidents/2025-10/Incidenten 202510.xlsm")
    (def opt {:n-header-rows 2 :sheet-filter ".*" :values :values :loglevel :info :table "auto"})
    ;; (excel->db opt db excel-path)
    (def workbook (load-workbook excel-path))
    (def sheet "Incidenten")
    (def sheet-name sheet)
    (def rows (read-sheet opt workbook sheet-name (:n-header-rows opt)))

    (def sheet (.getSheet workbook sheet-name))
    (def rows-all       (->> sheet (.iterator) iterator-seq))
    (def rows-data      (drop n-header-rows rows-all))
    (def n-header-rows 2)
    (def rows-header    (take n-header-rows rows-all))
    (def headers        (combine-row-headers rows-header))
    (def cell-fn        (cell-value-fn opt workbook))
    (def read-row-fn    (partial read-row cell-fn))
    (def data (map read-row-fn rows-data))
    ;; (def rows2 (map add-rownr rows (range (inc (:n-header-rows opt)) 1e8)))
    ;; (sheet->db opt db workbook sheet)

    (def res (vec (map (partial ordered-zipmap headers) data)))
    (def res2 (map (partial ordered-zipmap headers) data))
    (def res3 (ordered-zipmap headers (first data)))

    (def keys headers)
    (def vals (first data))
    (def zm (zipmap keys vals)) ;; this one gives wrong order.
    )
  )


(comment
  (def excel-file "/Users/nicodevreeze/Downloads/data.xlsx")
  (def workbook (load-workbook excel-file))
  (def data-orig (read-sheet {} workbook "Export" 2))
  (def sheet (.getSheet workbook "Export"))
  (def rows-all (->> sheet (.iterator) iterator-seq))
  (def rows-header    (take 2 rows-all))
  (def rows-data      (drop 2 rows-all))
  (def read-header-fn (partial read-row get-cell-string-value))

  (def header-rows (map read-row rows-header))

  ;; eerst fn die 2 rijen kan combineren, daarna kijken hoe het met 1 en 3 gaat.
  (apply map + [[1 2 3] [4 5 6]])
  (apply map str [[1 2 3] [4 5 6]])
  (apply map str header-rows)
  (apply map combine-headers header-rows)

  (combine-row-headers rows-header)

  #_(def data (read-sheet2 {} workbook "Export" 2))
  )

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
  (log/debugf (format "Loading workbook: %s" path))
  (doto (WorkbookFactory/create (io/input-stream path))
    (.setMissingCellPolicy Row$MissingCellPolicy/CREATE_NULL_AS_BLANK)))
