/* eslint-disable no-undef */

import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-64.png";
import "../../assets/icon-80.png";
import "../../assets/icon-128.png";

import { getMode } from "./services/outlook-mail.js";
import { initTabs } from "./components/tab-manager.js";
import {
  populateTeamDropdown,
  populateWorkItemTypes,
  populateTypeFilter,
  showStatus,
} from "./components/ui-helpers.js";
import { init as initCreate } from "./components/create-ticket.js";
import { init as initSearch } from "./components/search-items.js";
import { init as initInsert } from "./components/insert-card.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";

    const mode = getMode();

    // Populate config-driven UI
    populateWorkItemTypes("work-item-types");
    populateTeamDropdown("team-select");
    populateTypeFilter("search-type-filter");

    // Initialize tabs
    initTabs(mode);

    // Initialize modules
    if (mode === "read") {
      initCreate();
    }
    initSearch();
    if (mode === "compose") {
      initInsert();
    }
  }
});
