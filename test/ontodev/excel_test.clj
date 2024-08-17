(ns ontodev.excel-test
  (:require [clojure.test :refer :all]
            [midje.sweet :refer :all]
            [ontodev.excel :refer :all]))

(defn check-row
  [row]
  (facts "check-row"
         (:integer row) => "1001"
         (:float row) => "1001.01"
         (:formula row) => "2002.01"))

(let [workbook (load-workbook "resources/test.xlsx")
      data     (read-sheet workbook)]
  (doall (map check-row data))
  (fact "sheet names" (list-sheets workbook) => (just ["Sheet1" "Foo" "Bar"]))
  (fact "sheet headers" (sheet-headers workbook "Sheet1") => (just ["Format" "Integer" "Float" "Formula"])))

;; wanted to merge this test with the one above, but get failures then.
(let [workbook (load-workbook "resources/test.xlsx")
      data     (read-sheet {:values :strings} workbook)]
  (doall (map check-row data)))

(defn check-row-values
  [row]
  (facts "check-row-values"
    (:integer row) => 1001 
    (:float row) => 1001.01
    (:formula row) => 2002.01))

(let [workbook (load-workbook "resources/test.xlsx")
      data     (read-sheet {:values :values} workbook)]
  (doall (map check-row-values data)))

(defn check-row-formatted-3
  [row]
  (facts "check-row-formatted-3"
    (:integer row) => "1001.00"
    (:float row) => "1001.01"
    (:formula row) => "2002.01"))

(defn check-row-formatted-4
  [row]
  (facts "check-row-formatted-4"
    (:integer row) => "$1,001.00"
    (:float row) => "$1,001.01"
    (:formula row) => "$2,002.01"))

(defn check-row-formatted-5
  [row]
  (facts "check-row-formatted-5"
    (:integer row) => "1.00E+03"
    (:float row) => "1.00E+03"
    (:formula row) => "2.00E+03"))

(let [workbook (load-workbook "resources/test.xlsx")
      data     (read-sheet {:values :formatted} workbook)]
  (check-row-formatted-3 (nth data 3))
  (check-row-formatted-4 (nth data 4))
  (check-row-formatted-5 (nth data 5)))

(deftest show-version
  (testing "Showing current version"
    (println (format "Excel lib version: %s" (version)))
    (is (= (version) (version)))))

;; 2024-08-06: generate dummy error to prevent updating the POI lib. Replace with real test that shows the error.
(fact "dummy test wrt POI 5.3.0" "Dummy" => "ok")

;; 2024-08-10: lein ancient does not look at Midje results, so add a failing normal test.
(deftest a-test
  (testing "Failing wrt POI 5.3.0 issues"
    (is (= 0 1))))
