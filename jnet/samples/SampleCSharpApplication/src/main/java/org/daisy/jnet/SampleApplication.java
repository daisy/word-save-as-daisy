package org.daisy.jnet;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class SampleApplication {

	private String testString = "this is a sample string";
	
	public SampleApplication() {
		System.out.println(this.getClass().getCanonicalName() + " constructor called");
	}
	
	public SampleApplication(String testString) {
		System.out.println(this.getClass().getCanonicalName() + " constructor with argument called");
		this.testString = testString;
	}

	public void testMapOfListOfString(Map<String, List<String>> optionsTestMap) {
		List<String> test = new ArrayList<>();

		Integer.valueOf(4);
		for (Map.Entry<String,List<String>> entry:  optionsTestMap.entrySet()) {
			System.out.println(entry.getKey() + ": ");
			for (String subValue : entry.getValue()) {
				System.out.println(subValue);
			}
		}
	}

	public static void main(String[] args) {
		System.out.println("This is a simple execution test");
		int count = 0;
		for(String arg : args) {
			System.out.println("Argument " + (++count) + " - " + arg);
		}
		
		SampleApplication testing = new SampleApplication("This is a test string from static main");
		System.out.println(testing.getTestString());
		
	}
	
	public static void main() {
		System.out.println("This is a simple execution test with no args");
		
	}
	
	public String getTestString() {
		return this.testString;
	}
	
	public void setTestString(String message) {
		this.testString = message;
	}

}
