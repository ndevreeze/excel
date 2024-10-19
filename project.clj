(defproject ndevreeze/excel "0.3.2-SNAPSHOT"
  :description "A thin Clojure wrapper around a small part of Apache POI for
                reading .xlsx files."
  :url "http://github.com/ndevreeze/excel"
  :license {:name "Simplified BSD License"
            :url "http://opensource.org/licenses/BSD-2-Clause"}
  :dependencies [[org.clojure/clojure "1.12.0"]
                 [org.clojure/tools.logging "1.3.0"]
                 ;; 2024-08-17: fixed now, 5.2.5 not needed anymore.
                 [org.apache.poi/poi-ooxml "5.3.0"]
                 [clojure.java-time "1.4.2"]

                 ;; 2024-08-20: message: ERROR Log4j2 could not find a
                 ;; logging implementation. Please add log4j-core to
                 ;; the classpath. This indeed helps.
                 [org.apache.logging.log4j/log4j-api "2.24.1"]
                 [org.apache.logging.log4j/log4j-core "2.24.1"]]
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
