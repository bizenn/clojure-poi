(defproject org.visha/clojure-poi "0.0.1-SNAPSHOT"
  :description "Simple wrapper for clojure using Apache POI"
  :dependencies [[org.apache.poi/poi "3.8-beta4"]
                 [org.apache.poi/poi-ooxml "3.8-beta4"]]
  :dev-dependencies [[org.clojure/clojure "1.3.0"]
                     [swank-clojure "1.3.3"]]
  :source-path "src/main/clojure"
  :test-path "src/test/clojure"
  :resources-path "src/main/resources"
  :dev-resources-path "src/test/resources"
  :repl-init-script "src/main/script/swank-init.clj")
