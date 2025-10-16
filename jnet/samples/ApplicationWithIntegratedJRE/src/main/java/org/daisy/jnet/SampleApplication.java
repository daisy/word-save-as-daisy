package org.daisy.jnet;

public class SampleApplication {

	private String testString = "this is a sample string";
	
	public SampleApplication() {
		System.out.println(this.getClass().getCanonicalName() + " constructor called");
	}
	
	public SampleApplication(String testString) {
		System.out.println(this.getClass().getCanonicalName() + " constructor with argument called");
		this.testString = testString;
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
	
	public String getTestString() {
		return this.testString;
	}
	
	public void setTestString(String message) {
		this.testString = message;
	}

}
