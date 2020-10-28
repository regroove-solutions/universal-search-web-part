import * as React from "react";
import styles from "./SettingsPanel.module.scss";
import { Dropdown } from "office-ui-fabric-react/lib/Dropdown";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Slider } from "office-ui-fabric-react/lib/Slider";
import { PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import {
  defaultSearchTargets,
  defaultSettings
} from "../../../shared/constants";
import { Settings, SearchTarget } from "../../../shared/types";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import {
  UniversalSearch,
  IUniversalSearchProps
} from "../../universalSearch/components/UniversalSearch";
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";
import Reorder, { reorder } from "react-reorder";
import { IconButton } from "office-ui-fabric-react/lib/Button";

export interface ISettingsPanelProps {
  getSettings: () => Promise<Settings>;
  updateSettings: (val: Settings) => Promise<void>;
}

export const SettingsPanel: React.FC<ISettingsPanelProps> = ({
  getSettings,
  updateSettings
}: ISettingsPanelProps) => {
  const [settings, setSettings] = React.useState<Settings>();
  const [actualTheme, setTheme] = React.useState("");
  const [isLoading, setIsLoading] = React.useState(false);

  const getStoredSettings = async (): Promise<Settings> => {
    try {
      return {
        ...defaultSettings,
        ...(await getSettings())
      };
    } catch (e) {
      return defaultSettings;
    }
  };

  React.useEffect(() => {
    getStoredSettings().then((res) => setSettings(res));
  }, []);

  const onTextFieldChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setSettings({ ...settings, [e.target.name]: e.target.value });
  };

  const onDropdownChange = (
    e: React.FormEvent,
    value: SearchTarget & { selected: boolean }
  ) => {
    setSettings({
      ...settings,
      searchTargetKeys: value.selected
        ? [...settings.searchTargetKeys, value.key]
        : settings.searchTargetKeys.filter((k) => k !== value.key)
    });
  };

  const onSliderChange = (key: string, value: number) => {
    setSettings({ ...settings, [key]: value });
  };

  const onToggleChange = (
    e: React.MouseEvent<HTMLElement>,
    checked: boolean
  ) => {
    setSettings({ ...settings, showLogo: checked });
  };

  const onReorder = (
    event: React.MouseEvent<any>,
    prevIndex: number,
    nextIndex: number
  ) => {
    if (nextIndex < 0 || nextIndex >= settings.searchTargetKeys.length) return;
    setSettings({
      ...settings,
      searchTargetKeys: reorder(settings.searchTargetKeys, prevIndex, nextIndex)
    });
  };

  const validateURL = (value: string): string => {
    if (!settings.searchTargetKeys.filter((t) => t === "custom").length)
      return "";
    return value.match(/http(s)?:\/\/[\w.]+\.\w+/)
      ? ""
      : "Enter a search URL, e.g. 'https://google.ca/search?q='";
  };

  const validateHex = (value: string): string => {
    !value || !value.match(/#([A-F0-9]{3}){1,2}$/i)
      ? setTheme("#464775")
      : setTheme("");
    return value && !value.match(/#([A-F0-9]{3}){1,2}$/i)
      ? "Enter a valid hex colour code, e.g. '#F0F0F0'"
      : "";
  };

  const saveSettings = (newSettings: Settings = settings) => {
    setIsLoading(true);
    let searchTargets: SearchTarget[];
    if (
      !newSettings.customSearchUrl ||
      validateURL(newSettings.customSearchUrl)
    ) {
      searchTargets = newSettings.searchTargets.filter(
        (t) => t.key !== "custom"
      );
    } else {
      searchTargets = newSettings.searchTargets.map((t) =>
        t.key === "custom"
          ? {
              ...t,
              baseUrl: newSettings.customSearchUrl,
              text: newSettings.customSearchName || "Custom"
            }
          : t
      );
    }
    const themeColour = actualTheme || newSettings.themeColour;
    const updatedSettings: Settings = {
      ...newSettings,
      searchTargets: searchTargets,
      themeColour: themeColour
    };
    updateSettings(updatedSettings).then(() => setIsLoading(false));
  };

  const getPreviewSearchTargets = (): SearchTarget[] => {
    return settings.searchTargetKeys.map((k) => {
      if (k === "custom") {
        return {
          ...defaultSearchTargets.custom,
          text: settings.customSearchName || "Custom"
        };
      }
      return defaultSearchTargets[k];
    });
  };

  if (!settings || isLoading)
    return (
      <div className={styles.spinnerContainer}>
        <Spinner size={SpinnerSize.large} />
      </div>
    );

  return (
    <>
      <div className={styles.preview}>
        <UniversalSearch
          {...(settings as IUniversalSearchProps)}
          searchTargets={getPreviewSearchTargets()}
          disabled
          themeColour={actualTheme || settings.themeColour}
        />
      </div>
      <form className={styles.settingsPanel} onSubmit={() => saveSettings()}>
        <div>
          <Fabric className={styles.fabricText}>Search Options</Fabric>
          <Dropdown
            onChange={onDropdownChange}
            defaultSelectedKeys={settings.searchTargetKeys}
            label="Search targets"
            options={(Object as any).values(defaultSearchTargets)}
            multiSelect
            calloutProps={{ className: styles.dropdown }}
          />
          <Reorder
            reorderId="targets"
            draggedClassName="dragged"
            lock="horizontal"
            component="ul"
            onReorder={onReorder}
            holdTime={200}
            className={styles.reorder}
          >
            {settings.searchTargetKeys.map((item, index) => (
              <li key={index} id={item} className="">
                {defaultSearchTargets[item].text}
                <IconButton
                  iconProps={{ iconName: "ChevronUp" }}
                  onClick={(e) => onReorder(e, index, index - 1)}
                />
                <IconButton
                  iconProps={{ iconName: "ChevronDown" }}
                  onClick={(e) => onReorder(e, index, index + 1)}
                />
              </li>
            ))}
          </Reorder>
          <TextField
            name="customSearchName"
            onChange={onTextFieldChange}
            defaultValue={settings.customSearchName}
            label="Custom search title"
          />
          <TextField
            onGetErrorMessage={validateURL}
            name="customSearchUrl"
            onChange={onTextFieldChange}
            defaultValue={settings.customSearchUrl}
            label="Custom search URL"
          />
        </div>
        <div>
          <Fabric className={styles.fabricText}>Style Options</Fabric>
          <Slider
            onChange={(value) => onSliderChange("boxWidth", value)}
            value={settings.boxWidth}
            label="Search box width (%)"
            min={32}
            max={100}
            step={4}
            disabled={false}
          />
          <Slider
            onChange={(value) => onSliderChange("boxHeight", value)}
            value={settings.boxHeight}
            label="Search box height (px)"
            min={28}
            max={60}
            step={4}
          />
          <Slider
            onChange={(value) => onSliderChange("borderWidth", value)}
            value={settings.borderWidth}
            label="Border width (px)"
            min={0}
            max={8}
          />
          <TextField
            onChange={onTextFieldChange}
            onGetErrorMessage={validateHex}
            name="themeColour"
            defaultValue={settings.themeColour}
            label="Theme colour"
          />
          <Toggle
            className={styles.toggle}
            label="Show Navo Logo"
            onChange={onToggleChange}
            defaultChecked={settings.showLogo}
          />
        </div>
      </form>
      <PrimaryButton className={styles.submit} onClick={() => saveSettings()}>
        Save Settings
      </PrimaryButton>
    </>
  );
};
