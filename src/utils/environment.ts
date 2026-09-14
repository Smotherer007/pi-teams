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
 * those sessions get the device code flow instead.
 *
 * Set `PI_TEAMS_NO_BROWSER=1` to force that fallback anywhere.
 */
export function canOpenBrowser(): boolean {
	if (process.env.PI_TEAMS_NO_BROWSER) return false;

	// A remote shell: the browser would appear on the wrong machine.
	if (process.env.SSH_CONNECTION || process.env.SSH_TTY || process.env.SSH_CLIENT) return false;

	if (process.platform === "darwin" || process.platform === "win32") return true;

	// Linux and the BSDs need a display server for a browser to appear at all.
	return !!(process.env.DISPLAY || process.env.WAYLAND_DISPLAY);
}

/** Why the interactive flow is unavailable, phrased for a user-facing message. */
export function browserUnavailableReason(): string {
	if (process.env.PI_TEAMS_NO_BROWSER) return "PI_TEAMS_NO_BROWSER is set";
	if (process.env.SSH_CONNECTION || process.env.SSH_TTY || process.env.SSH_CLIENT) {
		return "this is an SSH session, so the browser would open on the wrong machine";
	}
	if (process.platform !== "darwin" && process.platform !== "win32") {
		return "no display server is available (DISPLAY and WAYLAND_DISPLAY are unset)";
	}
	return "no browser is available";
}
