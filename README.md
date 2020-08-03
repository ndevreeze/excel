# ontodev.excel / ndevreeze.excel 

A thin [Clojure](http://clojure.org) wrapper around a small part of [Apache POI](http://poi.apache.org) for reading `.xlsx` files. 

For a more complete implementation, see the `incanter-excel` module from the [Incanter](https://github.com/liebke/incanter) project. Althought, currently (August 2020) Incanter uses POI 3.16.

## Installation

Leiningen/Boot

    [ndevreeze/excel "0.3.1"]

Clojure CLI/deps.edn

    ndevreeze/excel {:mvn/version "0.3.1"}

[![Clojars Project](https://img.shields.io/clojars/v/ndevreeze/excel.svg)](https://clojars.org/ndevreeze/excel)

## Usage

`require` the namespace:

    (ns your.project
      (:require [ontodev.excel :as xls]))

Use it to load a workbook and read sheets:

    (let [workbook (xls/load-workbook "test.xlsx")
          sheet    (xls/read-sheet workbook "Sheet1")]
      (println "Sheet1:" (count sheet) (first sheet)))

Use the options parameter to get data back in different data formats:
    (xls/read-sheet {:values :values} workbook "Sheet1")
    (xls/read-sheet {:values :strings} workbook "Sheet1")
    (xls/read-sheet {:values :formatted} workbook "Sheet1")

## Related projects

Thanks to James Overton (ontodev) for starting this library.

* https://github.com/ontodev/excel - at version 0.2.5, using POI 3.8
* https://github.com/joshuaeckroth/excel - another fork, version 0.3.0, using POI 4.0.1

This version adds option to get data in the correct datatypes, including dates and times. Also there is an option to get the formatted data.

## Testing

    $ lein midje
    nil
    All checks (83) succeeded.

## History

* 0.2.5 - base version from ontodev
* 0.3.0 - another fork by joshuaeckroth
* 0.3.1 - use POI 4.1.2, return different data formats.

## License

Copyright © 2014, James A. Overton; Copyright © 2020, Nico de Vreeze

Distributed under the Simplified BSD License: [http://opensource.org/licenses/BSD-2-Clause]()

