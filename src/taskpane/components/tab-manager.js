/* Tab navigation manager */

const TABS = ["create", "search", "insert"];

export function initTabs(mode) {
  const nav = document.getElementById("tab-nav");
  const buttons = nav.querySelectorAll(".tab-btn");

  buttons.forEach((btn) => {
    const tab = btn.dataset.tab;

    // Disable tabs based on mode
    if (mode === "compose" && tab === "create") {
      btn.disabled = true;
      btn.classList.remove("active");
    }
    if (mode === "read" && tab === "insert") {
      btn.disabled = true;
      btn.classList.remove("active");
    }

    btn.addEventListener("click", () => {
      if (btn.disabled) return;
      showTab(tab);
    });
  });

  // Show the default tab based on mode
  const defaultTab = mode === "compose" ? "search" : "create";
  showTab(defaultTab);
  nav.style.display = "flex";
}

export function showTab(tabName) {
  const buttons = document.querySelectorAll(".tab-btn");
  buttons.forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.tab === tabName);
  });

  TABS.forEach((t) => {
    const panel = document.getElementById(`tab-${t}`);
    if (panel) panel.style.display = t === tabName ? "block" : "none";
  });
}
