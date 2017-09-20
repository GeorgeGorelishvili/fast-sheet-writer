package org.george.gorelishvili.export.common;

public class OSDetection {
 
	private static String OS = System.getProperty("os.name").toLowerCase();
 
	static boolean isWindows() {
		return isOs("win");
	}
 
	public static boolean isMac() {
		return isOs("mac");
	}
 
	public static boolean isUnix() {
		return (isOs("nix") || isOs("nux") || isOs("aix"));
	}
 
	public static boolean isSolaris() {
		return isOs("sunos");
	}

	private static boolean isOs(String osName) {
		return OS.contains(osName);
	}
}