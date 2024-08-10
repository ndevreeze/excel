(defproject ndevreeze/excel "0.3.2-SNAPSHOT"
  :description "A thin Clojure wrapper around a small part of Apache POI for
                reading .xlsx files."
  :url "http://github.com/ndevreeze/excel"
  :license {:name "Simplified BSD License"
            :url "http://opensource.org/licenses/BSD-2-Clause"}
  :dependencies [[org.clojure/clojure "1.11.4"]
                 [org.clojure/tools.logging "1.3.0"]
                 ;; [org.apache.poi/poi-ooxml "5.3.0"]
                 ;; 2024-08-01: 5.3.0 gives errors, 5.2.5 is still ok.
                 [org.apache.poi/poi-ooxml "5.2.5"]
                 [clojure.java-time "1.4.2"]]
  :profiles
  {:dev {:dependencies [[midje "1.10.10"]
                        [lazytest "1.2.3"]]
         :plugins [[lein-midje "3.2.2"]
                   [lein-marginalia "0.7.1"]]}}

  :codox
  {:output-path "docs/api"
   :metadata {:doc/format :markdown}
   :source-uri "https://github.com/ndevreeze/excel/blob/master/{filepath}#L{line}"}

  :repositories [["releases" {:url "https://clojars.org/repo/"
                              :creds :gpg}]]

  )
