# Slack App Readme

This Slack app is designed to perform bulk operations such as sending invitations and updating user display names based on an Excel spreadsheet. It utilizes the Slack API and several Python libraries to interact with Slack workspaces.

## Prerequisites

Before using this app, ensure that you have the following:

- Python 3.x installed on your machine.
- The required Python libraries installed. You can install them using `pip`:
  - slack_bolt
  - dotenv
  - openpyxl
  - pandas

## Configuration

1. Clone or download the source code of this Slack app.
2. Create a new Slack app in your workspace through the Slack API website.
3. Obtain the following tokens from your Slack app:
   - `SLACK_BOT_TOKEN`: Bot token for your app.
   - `SLACK_APP_TOKEN`: App token for your app.
   - `SLACK_USER_TOKEN`: User token for your app.
4. Create a `.env` file in the root directory of the app and populate it with the obtained tokens:
   ```
   SLACK_WORKSPACE=<your-slack-workspace>
   SLACK_BOT_TOKEN=<your-slack-bot-token>
   SLACK_APP_TOKEN=<your-slack-app-token>
   SLACK_USER_TOKEN=<your-slack-user-token>
   ```
5. Place the Excel spreadsheet containing the access list in the same directory as the app and update the `EXCEL_FILE` constant in the code with the correct filename.

## Usage

To run the Slack app, execute the following command in the terminal:

```shell
python <filename>.py
```

Replace `<filename>` with the name of the Python script file containing the Slack app code.

### Sending Invitations

To send bulk invitations to Slack channels, use the `/invite_bulk` command in your Slack workspace. The app will read the specified Excel sheet(s) and send invitations to the corresponding email addresses. The invited users will be added to specific channels based on matching criteria.

Usage: `/invite_bulk [sheet1, sheet2, ...]`

- If no sheet names are provided, invitations will be sent to all sheets in the Excel file.
- If sheet names are specified, invitations will be sent only to those sheets present in the Excel file.

### Updating User Display Names

To update user display names in bulk, use the `/set_display_name` command in your Slack workspace. The app can update user profiles based on email addresses or read the entire Excel sheet to update display names.

Usage: `/set_display_name [email, fullname, displayname]`

- If email, fullname, and displayname are provided as parameters, the app will update the display name of the specified user.
- If no parameters are provided, the app will read the Excel sheet and update the display names for all users listed.

## Debug Mode

By default, the app runs in normal mode, with debug mode disabled. However, you can enable debug mode by setting `DEBUG=True` in the code. This will display additional information during execution, such as API responses and user updates.

Note: It is recommended to disable debug mode in a production environment.

## Contributions

Contributions to this Slack app are welcome. Feel free to submit bug reports, feature requests, or pull requests to enhance its functionality.

## License

This project is licensed under the [MIT License](https://opensource.org/licenses/MIT).