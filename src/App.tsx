import React from "react";
import { RouterHelper, TeamsThemeHelper, AuthHelper } from "./helpers";
import { ThemePrepared } from "@fluentui/react-northstar";
import * as msTeams from "@microsoft/teams-js";
import { ThemeProvider } from "react-bootstrap";

export default class App extends React.Component<IAppProps, IAppState> {
  constructor(props: IAppProps) {
    super(props);

    this.state = {
      theme: TeamsThemeHelper.getTheme("default"),
      loggedIn: AuthHelper.IsUserLoggedIn(),
    };

    msTeams.app.initialize();
    msTeams.app.registerOnThemeChangeHandler(this.updateTheme.bind(this));
    msTeams.app.getContext().then((context) => {
      this.updateTheme(context.app.theme);
    });
  }

  render() {
    return (
      <ThemeProvider theme={this.state.theme}>
        {this.state.loggedIn ? (
          <RouterHelper.AuthenticatedRoutes />
        ) : (
          <RouterHelper.UnauthenticatedRoutes />
        )}
      </ThemeProvider>
    );
  }

  private updateTheme(themeString: string | undefined): void {
    this.setState({
      theme: TeamsThemeHelper.getTheme(themeString),
    });
  }
}

interface IAppProps {}

interface IAppState {
  theme: ThemePrepared;
  loggedIn: boolean;
}
