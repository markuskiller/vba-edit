/* Runs before first paint so a stored theme choice does not flash. */
try {
  var storedTheme = localStorage.getItem("vba-edit-theme");
  if (storedTheme === "light" || storedTheme === "dark") document.documentElement.dataset.theme = storedTheme;
} catch (e) { /* storage unavailable: follow the system setting */ }
