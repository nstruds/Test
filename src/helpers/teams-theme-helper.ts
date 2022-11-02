import {
  teamsDarkTheme,
  teamsHighContrastTheme,
  teamsTheme,
  ThemePrepared,
} from "@fluentui/react-northstar";
import { TeamsThemes } from "../constants";

export default class TeamsThemeHelper {
  public static getTheme(themeStr: string | undefined): ThemePrepared {
    themeStr = themeStr || "";

    switch (themeStr) {
      case TeamsThemes.dark:
        return teamsDarkTheme;
      case TeamsThemes.contrast:
        return teamsHighContrastTheme;
      case TeamsThemes.default:
      default:
        return teamsTheme;
    }
  }
}
