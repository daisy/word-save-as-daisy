# SaveAsDAISY conversion library

Library holding the code to convert a word document to dtbook XML, and then to other formats using DAISY Pipeline 1 and 2.



## Building the DAISY Pipeline 2 and word-to-dtbook module for SaveAsDAISY

To update the DAISY Pipeline 2, you will need to retrieve the latest version of the pipeline-assembly project with the "SimpleAPI" bridge to run the pipeline using JNI.

On windows, you can build the pipeline with the minimum of interfaces using the following maven command : 

```shell
mvn clean install -Passemble-word-addin-dir -Punpack-updater-win -Punpack-updater-gui-win -Pcompile-simple-api -Pwithout-webservice -Pwithout-gui -Pwithout-persistence -Pwithout-updater -Pwithout-osgi  -Pcopy-artifacts 
```

As we have started to migrate the word-to-dtbook xsl conversion code to XSLT2, the xslt is now migrated to a specific DAISY pipeline 2 module project, that is placed in the "Lib\pipeline2-module" folder (for distribution purpose)

Go to this folder and simply launch the command to build the module
```
mvn clean install
```

You can then either copy the resulting jar and the target/dependencies/poi-ooxml-5.0.0.jar file to the daisy pipeline 2 folder for tests or just build the full plugin including its package (the jar mentioned should be referenced in the wix project so you won't have anything to do for it to be included in the distributed pipeline 2)

