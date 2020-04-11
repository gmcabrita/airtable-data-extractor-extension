import Papa from "papaparse";
import "../img/icon-128.png";
import "../img/icon-34.png";

let tabData = {};

function process(tab, type = "json") {
  if (tabData[tab.id] && tabData[tab.id].json != null) {
    generateDownload(tabData[tab.id].json, type);
  } else if (tabData[tab.id] && tabData[tab.id].fn != null) {
    tabData[tab.id]
      .fn()
      .then((response) => response.json())
      .then((json) => {
        tabData[tab.id].json = json;
        generateDownload(tabData[tab.id].json, type);
      })
      .catch((error) => {
        console.error(error);
      });
  }
}

chrome.contextMenus.create({
  title: "Download JSON",
  contexts: ["page", "browser_action"],
  documentUrlPatterns: ["*://airtable.com/embed/shr*", "*://airtable.com/shr*"],
  onclick: (event, tab) => {
    process(tab, "json");
  },
});

chrome.contextMenus.create({
  title: "Download CSV",
  contexts: ["page", "browser_action"],
  documentUrlPatterns: ["*://airtable.com/embed/shr*", "*://airtable.com/shr*"],
  onclick: (event, tab) => {
    process(tab, "csv");
  },
});

function generateDownload(json, type) {
  const columns = {};
  json.data.table.columns.forEach((column) => {
    columns[column.id] = column;
  });

  const result = json.data.table.rows.map((row) => {
    const rowData = {};
    for (let [key, value] of Object.entries(row.cellValuesByColumnId)) {
      const column = columns[key];

      // TODO: convert this into a switch case
      if (column.type == "multiSelect") {
        rowData[column.name] = [value]
          .flat()
          .map((choice) => column.typeOptions.choices[choice].name);
      } else if (column.type == "multipleAttachment") {
        rowData[column.name] = value.map((attachment) => attachment.url);
      } else if (column.type == "foreignKey") {
        rowData[column.name] = value.map((fk) => fk.foreignRowDisplayName);
      } else if (column.type == "select") {
        rowData[column.name] = column.typeOptions.choices[value].name;
      } else if (column.type == "richText") {
        rowData[column.name] = value.documentValue
          .map((e) => e.insert)
          .join(" ");
      } else {
        rowData[column.name] = value;
      }
    }

    return rowData;
  });

  let blob;
  let filename;
  if (type == "json") {
    blob = new Blob([JSON.stringify(result)], { type: "application/json" });
    filename = `${new Date().getTime()}.json`;
  } else if (type == "csv") {
    blob = new Blob([Papa.unparse(result)], {
      type: "text/csv;charset=utf-8;",
    });
    filename = `${new Date().getTime()}.csv`;
  }

  if (blob) {
    const url = URL.createObjectURL(blob);
    chrome.downloads.download({
      url: url,
      filename: filename,
    });
  }
}

chrome.tabs.onRemoved.addListener((tabId, removeInfo) => {
  delete tabData[tabId];
});

chrome.webRequest.onBeforeSendHeaders.addListener(
  async (details) => {
    let queryParams = new URLSearchParams(details.url.split("?")[1]);
    if (!queryParams.has("z")) {
      let headers = details.requestHeaders.reduce((acc, { name, value }) => {
        acc[name] = value;
        return acc;
      }, {});

      tabData[details.tabId] = {
        fn: () => {
          return fetch(`${details.url}&z=1`, {
            credentials: "include",
            headers: headers,
          });
        },
      };
    }
  },
  {
    urls: ["*://airtable.com/*/view/*/readSharedViewData*"],
  },
  ["extraHeaders", "requestHeaders"]
);
