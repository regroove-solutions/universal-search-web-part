import * as React from "react";
import { SearchTarget } from "../../../shared/types";
import {
  DuckDuckGoLogoSVG,
  GoogleLogoSVG,
  NavoLogoSVG
} from "../../../shared/svg";
import {
  ISearchBoxStyles,
  SearchBox
} from "office-ui-fabric-react/lib/SearchBox";
import styles from "./UniversalSearch.module.scss";
import { IButtonProps, PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { registerIcons } from "office-ui-fabric-react/lib/Styling";
import {
  MessageBar,
  MessageBarType
} from "office-ui-fabric-react/lib/MessageBar";
import { Fabric } from "office-ui-fabric-react/lib/Fabric";
import {
  IContextualMenuItem,
  IContextualMenuProps
} from "office-ui-fabric-react/lib/ContextualMenu";
import { initializeIcons } from "office-ui-fabric-react/lib/Icons";

initializeIcons();
registerIcons({
  icons: {
    DuckDuckGoLogo: DuckDuckGoLogoSVG,
    GoogleLogo: GoogleLogoSVG
  }
});

export interface IUniversalSearchProps {
  instanceId: string;
  boxWidth: number;
  boxHeight: number;
  newTab: boolean;
  themeColour: string;
  searchTargets: SearchTarget[];
  borderWidth: number;
  usePreference: boolean;
  showLogo: boolean;
  disabled?: boolean;
}

export const UniversalSearch: React.FC<IUniversalSearchProps> = ({
  instanceId,
  boxWidth,
  newTab,
  themeColour,
  searchTargets,
  boxHeight,
  borderWidth,
  usePreference,
  showLogo,
  disabled
}: IUniversalSearchProps) => {
  const getStoredTarget = (): SearchTarget => {
    if (!usePreference) return searchTargets[0];
    try {
      const storedTarget: SearchTarget = JSON.parse(
        localStorage.getItem("searchTarget")
      )[instanceId];
      if (searchTargets.some((target) => target.key === storedTarget.key))
        return storedTarget;
      return searchTargets[0];
    } catch (e) {
      return searchTargets[0];
    }
  };

  const [query, setQuery] = React.useState();
  const [chosenTarget, setChosenTarget] = React.useState(getStoredTarget);

  React.useEffect(() => {
    if (!searchTargets.some((target) => target.key === chosenTarget.key))
      setChosenTarget(searchTargets[0]);
  }, [searchTargets]);

  const handleSearch = (input: string, url?: string) => {
    if (!input || disabled) return;

    const replaceValues: [RegExp, string][] = [
      [/\//g, "%2F"],
      [/\?/g, "%3F"],
      [/&/g, "%26"],
      [/=/g, "%3D"],
      [/\+/g, "%2B"],
      [/#/g, "%23"]
    ];

    let URI = encodeURI(input);
    replaceValues.forEach((pair) => {
      const [regex, newValue] = pair;
      URI = URI.replace(regex, newValue);
    });
    URI = URI.replace(/%20/g, "+");
    window.open(
      `${url || chosenTarget.baseUrl}${URI}`,
      newTab ? "_blank" : "_self"
    );
  };

  const handleChangeTarget = (target: SearchTarget) => {
    setChosenTarget(target);
    let savedTargets: { [key: string]: SearchTarget } = {};
    try {
      savedTargets = JSON.parse(localStorage.getItem("searchTarget")) || {};
    } catch (e) {}
    savedTargets[instanceId] = target;
    localStorage.setItem("searchTarget", JSON.stringify(savedTargets));
    handleSearch(query, target.baseUrl);
  };

  if (!chosenTarget)
    return (
      <MessageBar messageBarType={MessageBarType.warning}>
        Please select a search target in the property pane or settings tab
      </MessageBar>
    );

  const menuProps: IContextualMenuProps = {
    calloutProps: { className: styles.searchDropdown },
    items: searchTargets.map(
      (target): IContextualMenuItem => ({
        key: target.key,
        text: target.text,
        iconProps: {
          iconName: target.iconName,
          style: { color: themeColour, fill: themeColour }
        },
        onClick: () => handleChangeTarget(target)
      })
    )
  };
  const fontSize = boxHeight / 8 + 12;
  const boxStyle: React.CSSProperties = {
    height: boxHeight,
    padding: borderWidth,
    backgroundColor: themeColour,
    fontSize: fontSize
  };
  const searchBoxStyles: Partial<ISearchBoxStyles> = {
    iconContainer: {
      color: `${themeColour} !important`,
      width: fontSize * 2,
      fontSize: fontSize
    }
  };
  const buttonStyleProps: Partial<IButtonProps> = {
    style: {
      backgroundColor: themeColour,
      fontSize: fontSize,
      minWidth: fontSize * 3
    },
    styles: { menuIcon: { fontSize: fontSize }, icon: { fontSize: fontSize } },
    iconProps: { iconName: chosenTarget.iconName, style: { width: fontSize } }
  };

  return (
    <div style={{ width: boxWidth + "%" }} className={styles.searchWebPart}>
      <div className={styles.universalSearch} style={boxStyle}>
        <SearchBox
          className={styles.searchBox}
          placeholder="Search"
          onSearch={handleSearch}
          iconProps={{ style: { opacity: 100 } }}
          onChange={setQuery}
          style={{ fontSize: boxHeight / 8 + 12 }}
          styles={searchBoxStyles}
        />
        {searchTargets.length > 1 ? (
          <PrimaryButton
            className={styles.targetButton}
            menuProps={menuProps}
            {...buttonStyleProps}
          />
        ) : (
          <PrimaryButton
            onClick={() => handleSearch(query)}
            style={buttonStyleProps.style}
          >
            🡢
          </PrimaryButton>
        )}
      </div>
      {showLogo && (
        <div className={styles.poweredBy}>
          <Fabric>Powered by </Fabric>
          <a href="https://getnavo.com" target="_blank">
            {NavoLogoSVG}
          </a>
        </div>
      )}
    </div>
  );
};

export const UniversalSearchTeams: React.FC<IUniversalSearchProps> = (
  props: IUniversalSearchProps
) => (
  <div className={styles.teamsContainer}>
    <UniversalSearch {...props} />
  </div>
);
