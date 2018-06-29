
import static java.lang.System.out;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.concurrent.TimeoutException;

// Run example
// C:\_qtp\resources\JavaProg>java RunUftTest "C:\_qtp\tests\_Ne opredelen\S.001 Otkritie glavnih modulei" CAO Test

public class RunUftTest {

	public static void main(String[] args) {
		// Args: TESTPATH CAO TEST
		
		String testPath = args[0];
		String district = args[1];
		// TODO: remove '-' in uDeploy list of environment variable
		String environment = args[2].replace("-", "");
		
		System.out.println("{ testPath: "+ testPath + ";\r\n  district: " + district + ";\r\n  environment: " + environment + " }\r\n");
		
		/*String runCommand = "cscript.exe clocker.vbs";
		String execDirectory = "C:\\Temp";
		long timeout = 13000;//*/
		String runCommand = "cscript.exe RunTest.vbs \"" + testPath + "\" " + district + "-" + environment;
		String execDirectory = "C:\\_qtp\\resources\\VBscripts";
		long timeout = 900000; //*/  // 15 minutes for each test execution
		ExecProcessInfo execProcessInfo;
		int trialCount = 0;
		int maxTrialCount = 3;
		
		do {
			System.out.println("\r\n--------------------------------------------------------------");
			System.out.println("  TEST RUN #" + trialCount);
			execProcessInfo = execProcess(runCommand, execDirectory, timeout);
			System.out.println("  RUN STATUS: " + execProcessInfo.getExitStatus());
			trialCount++;
		} while (execProcessInfo.getExitStatus() != 0 && trialCount < maxTrialCount);	// 0 - success
		
		System.out.println("\r\n==============================================================");
		
		if (execProcessInfo.getExitStatus() != 0) { // print start script trace if test fails
			System.out.println("RunTest.vbs - OUTPUT:\r\n");
			printBuffer(execProcessInfo.inputStream);
		} else {
			System.out.println("TEST PASSED!\r\n");
		}

		System.exit(execProcessInfo.getExitStatus());
		
	}
	
	private static class Worker extends Thread {
		private final Process process;
		private Integer exit;

		private Worker(Process process) {
			this.process = process;
		}

		public void run() {
			try {
				exit = process.waitFor();
			} catch (InterruptedException ignore) {
				return;
			}
		}
	}
	
	private static class ExecProcessInfo {
		private Integer exitStatus;
		public BufferedReader inputStream;
		
		ExecProcessInfo() {
			setExitStatus(-1); // Fail status by default
			inputStream = null;
		}

		public Integer getExitStatus() { return exitStatus; }
		public void setExitStatus(Integer exitStatus) { this.exitStatus = exitStatus; }
	}
	
	
	private static void printBuffer(BufferedReader bufferedReader) {
		try {
			String s;
			while ((s = bufferedReader.readLine()) != null) {
				out.println(s);
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static ExecProcessInfo execProcess(String runCommand, String execDirectory, long timeout) {
		ExecProcessInfo execProcessInfo = new ExecProcessInfo();
		
		try {
			
			Process process = Runtime.getRuntime().exec(runCommand, null, new File(execDirectory));
			
			System.out.println("Initializing process  <" + runCommand + ">...");
			System.out.println("              in dir  <" + execDirectory + ">");
			System.out.println("        with timeout  " + timeout + "ms\r\n");
			execProcessInfo.inputStream = new BufferedReader(new InputStreamReader(process.getInputStream()));

			Worker worker = new Worker(process);
			worker.start();
			long startTime = 0;

			try {
				startTime = System.currentTimeMillis();
				worker.join(timeout);
				
				if (worker.exit != null) {	// process exited by himself
					System.out.println("         exit status  " + worker.exit);
					execProcessInfo.setExitStatus(worker.exit);
				}
				else {
					long procExecDuration = System.currentTimeMillis() - startTime;
					System.out.println("Process was interrupted after " + procExecDuration + "ms.");
					throw new TimeoutException();
				}
			} catch (InterruptedException ex) {
				worker.interrupt();
				long procExecDuration = System.currentTimeMillis() - startTime;
				System.out.println("InterruptedException. After " + procExecDuration + "ms of execution.");
				ex.printStackTrace();
				killUftProcess();
				killJP2LauncherProcess();
				killWerFaultProcess();
				Thread.currentThread().interrupt();
				throw ex;
			} catch (TimeoutException e) {
				killUftProcess();
				killJP2LauncherProcess();
				killWerFaultProcess();
			} finally {
				process.destroy();
			}

		} catch (IOException e) {
			e.printStackTrace();
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
		return execProcessInfo;
	}

	private static void killUftProcess() throws IOException {
		System.out.println("Killing <UFT.exe> process");
		Runtime.getRuntime().exec("taskkill /F /IM UFT.exe");
	}
	
	private static void killJP2LauncherProcess() throws IOException {
		System.out.println("Killing <jp2launcher.exe> process");
		Runtime.getRuntime().exec("taskkill /F /IM jp2launcher.exe");
	}
	
	private static void killWerFaultProcess() throws IOException {
		System.out.println("Killing <WerFault.exe> process");
		Runtime.getRuntime().exec("taskkill /F /IM WerFault.exe");
	}
}
