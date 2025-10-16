
public class DefaultPackageSample {

	public DefaultPackageSample() {
		System.out.println(this.getClass().getCanonicalName() + " constructor called");
	}
	
	public String getExecutablePath() {
		String javaExecutablePath = ProcessHandle.current()
			    .info()
			    .command()
			    .orElseThrow();
		return javaExecutablePath;
	}

	public static void main(String[] args) {
		System.out.println("main called");
	}


}
