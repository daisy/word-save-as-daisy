include engine/make/enable-java-shell.mk

.PHONY : default
default : resources/daisy-pipeline

ENGINE_VERSION := $(shell println(xpath(new File("engine/pom.xml"), "/*/*[local-name()='version']/text()"));)
zip_classifier := win

resources/daisy-pipeline : engine/target/assembly-$(ENGINE_VERSION)-$(zip_classifier).zip
	rm("$@");
	unzip(new File("$<"), new File("$(dir $@)"));

engine/target/assembly-$(ENGINE_VERSION)-$(zip_classifier).zip : \
		engine/pom.xml \
		$(shell Files.walk(Paths.get("engine/src")).filter(Files::isRegularFile).forEach(System.out::println);)
	exec("$(MAKE)", "-C", "engine", "zip-$(zip_classifier)",         \
	                                "--", "--without-osgi",          \
			                        "--with-simple-api",             \
			                        "--without-cli",                 \
			                        "--without-updater",             \
			                        "--without-webservice",          \
			                        "--without-persistence");

clean :
	rm("resources/daisy-pipeline");
	exec("$(MAKE)", "-C", "engine", "clean");


