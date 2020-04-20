import Papa from "papaparse";
import "../img/icon-128.png";
import "../img/icon-34.png";

let tabData = {};

async function process(tab, type = "json") {
  if (tabData[tab.id] && tabData[tab.id].fn) {
    const json = await tabData[tab.id].fn();
    tabData[tab.id].json = json;
    delete tabData[tab.id].fn;
  }

  if (tabData[tab.id] && tabData[tab.id].updateRowsFn) {
    const json = await tabData[tab.id].updateRowsFn();
    tabData[tab.id].json = {
      ...tabData[tab.id].json,
      rows: json.rows,
    };
    delete tabData[tab.id].updateRowsFn;
  }

  if (tabData[tab.id] && tabData[tab.id].json) {
    generateDownload(tabData[tab.id].json, type);
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
  json.columns.forEach((column) => {
    columns[column.id] = column;
  });

  let csvColumns = [];

  const result = json.rows.map((row) => {
    const rowData = {};

    for (let [key, value] of Object.entries(row.cellValuesByColumnId)) {
      const column = columns[key];

      if (!csvColumns.includes(column.name)) {
        csvColumns.push(column.name);
      }

      switch (column.type) {
        case "multiSelect":
          rowData[column.name] = [value]
            .flat()
            .map((choice) => column.typeOptions.choices[choice].name);
          break;
        case "multipleAttachment":
          rowData[column.name] = value.map((attachment) => attachment.url);
          break;
        case "foreignKey":
          rowData[column.name] = value.map((fk) => fk.foreignRowDisplayName);
          break;
        case "select":
          rowData[column.name] = column.typeOptions.choices[value].name;
          break;
        case "richText":
          rowData[column.name] = value.documentValue
            .map((e) => e.insert)
            .join(" ");
          break;
        default:
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
    blob = new Blob([Papa.unparse(result, { columns: csvColumns })], {
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
          })
            .then((response) => response.json())
            .then((json) => {
              return {
                rows: json.data.table.rows,
                columns: json.data.table.columns,
              };
            });
        },
      };
    }
  },
  {
    urls: ["*://airtable.com/*/view/*/readSharedViewData?*"],
  },
  ["extraHeaders", "requestHeaders"]
);

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
          })
            .then((response) => response.json())
            .then((json) => {
              return {
                rows: json.data.tableDatas[0].rows,
                // TODO: there might be id collisions in the future but this is ok for now
                columns: json.data.tableSchemas
                  .map((schema) => schema.columns)
                  .flat(),
              };
            });
        },
      };
    }
  },
  {
    urls: ["*://airtable.com/*/application/*/read?*"],
  },
  ["extraHeaders", "requestHeaders"]
);

chrome.webRequest.onBeforeSendHeaders.addListener(
  async (details) => {
    let queryParams = new URLSearchParams(details.url.split("?")[1]);
    if (!queryParams.has("z")) {
      let headers = details.requestHeaders.reduce((acc, { name, value }) => {
        acc[name] = value;
        return acc;
      }, {});

      tabData[details.tabId] = {
        ...tabData[details.tabId],
        updateRowsFn: () => {
          return fetch(`${details.url}&z=1`, {
            credentials: "include",
            headers: headers,
          })
            .then((response) => response.json())
            .then((json) => {
              return {
                rows: json.data.rows,
              };
            });
        },
      };
    }
  },
  {
    urls: ["*://airtable.com/*/table/*/readData?*"],
  },
  ["extraHeaders", "requestHeaders"]
);
