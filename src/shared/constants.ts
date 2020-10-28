import { SearchTarget, Settings } from "./types";

export const defaultSearchTargets: { [key: string]: SearchTarget } = {
  microsoft: {
    key: "microsoft",
    text: "Microsoft 365",
    iconName: "WindowsLogo",
    baseUrl: "https://www.office.com/search?auth=2&q="
  },
  google: {
    key: "google",
    text: "Google",
    iconName: "GoogleLogo",
    baseUrl: "https://www.google.com/search?q="
  },
  duckduckgo: {
    key: "duckduckgo",
    text: "DuckDuckGo",
    iconName: "DuckDuckGoLogo",
    baseUrl: "https://duckduckgo.com/?q="
  },
  bing: {
    key: "bing",
    text: "Bing",
    iconName: "BingLogo",
    baseUrl: "https://www.bing.com/search?q="
  },
  sharepoint: {
    key: "sharepoint",
    text: "SharePoint",
    iconName: "SharepointLogo",
    baseUrl: "/search/Pages/results.aspx?k="
  },
  custom: {
    key: "custom",
    text: "Custom",
    iconName: "Globe",
    baseUrl: ""
  }
};

export const defaultSettings: Settings = {
  instanceId: "preview-search",
  boxWidth: 44,
  boxHeight: 32,
  borderWidth: 2,
  customSearchName: "",
  customSearchUrl: "",
  themeColour: "#464775",
  searchTargetKeys: ["microsoft"],
  searchTargets: [
    {
      key: "microsoft",
      text: "Microsoft 365",
      iconName: "WindowsLogo",
      baseUrl: "https://www.office.com/search?auth=2&q="
    }
  ],
  usePreference: true,
  showLogo: true
};

export const webPartKey = "SearchWebPart";
