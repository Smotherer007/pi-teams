/**
 * Environment detection.
 *
 * Deliberately dependency-free: both the config layer and the auth layer need
 * to know whether a browser is reachable, and neither should have to import
 * the other to find out.
 */

/**
 * Can we realistically put a browser in front of this user?
 *
 * The interactive sign-in needs two things: a browser to open, and a loopback
 * listener that browser can reach. Over SSH both assumptions break — `open`
 * would launch a browser on the *remote* machine, where nobody is looking — so
 * those sessions get the device code flow instead. In WSL the browser is on
 * the Windows side of the same machine, which is what `isWsl` accounts for.
 *
 * Set `PI_TEAMS_NO_BROWSER=1` to force that fallback anywhere, or
 * `PI_TEAMS_BROWSER` to a command line that opens a URL where the built-in
 * launchers cannot.
 */
export function isWsl(): boolean {
	if (process.platform !== "linux") return false;
	return !!(process.env.WSL_DISTRO_NAME || process.env.WSL_INTEROP);
}

export function canOpenBrowser(): boolean {
	if (process.env.PI_TEAMS_NO_BROWSER) return false;

	// An explicit launcher is a statement that a browser is reachable, even where
	// the heuristics below would not find one — a WSL setup, or a machine whose
	// browser is reached through something custom.
	if (process.env.PI_TEAMS_BROWSER?.trim()) return true;

	// A remote shell: the browser would appear on the wrong machine.
	if (process.env.SSH_CONNECTION || process.env.SSH_TTY || process.env.SSH_CLIENT) return false;

	if (process.platform === "darwin" || process.platform === "win32") return true;

	// WSL is not a headless Linux: the browser lives on Windows and `openBrowser`
	// hands the URL over to it. DISPLAY says nothing either way — WSLg sets it on
	// Windows 11 and not on Windows 10, while the handoff works on both.
	if (isWsl()) return true;

	// Linux and the BSDs need a display server for a browser to appear at all.
	return !!(process.env.DISPLAY || process.env.WAYLAND_DISPLAY);
}

/** Why the interactive flow is unavailable, phrased for a user-facing message. */
export function browserUnavailableReason(): string {
	if (process.env.PI_TEAMS_NO_BROWSER) return "PI_TEAMS_NO_BROWSER is set";	if (process.env.SSH_CONNECTION || process.env.SSH_TTY || process.env.SSH_CLIENT) {
		return "this is an SSH session, so the browser would open on the wrong machine";
	}
	if (process.platform !== "darwin" && process.platform !== "win32") {
		return "no display server is available (DISPLAY and WAYLAND_DISPLAY are unset)";
	}
	return "no browser is available";
}
