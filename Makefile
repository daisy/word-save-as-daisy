include engine/assembly/make/enable-java-shell.mk

.PHONY : default
default : resources/daisy-pipeline

ENGINE_VERSION := $(shell println(xpath(new File("engine/pom.xml"), "/*/*[local-name()='version']/text()"));)
zip_classifier := win
resources/daisy-pipeline : engine/pipeline2-$(ENGINE_VERSION)_windows.zip
	rm("$@");
	unzip(new File("$<"), new File("$(dir $@)"));

engine/pipeline2-$(ENGINE_VERSION)_windows.zip : \
		engine/pom.xml
	exec("$(MAKE)", "-C", "engine", "dist-zip-$(zip_classifier)");

clean :
	rm("resources/daisy-pipeline");
	exec("$(MAKE)", "-C", "engine", "clean");


