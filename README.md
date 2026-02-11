# Schulportal Hessen Timetable Exporter

A robust Userscript to export timetables from the Schulportal Hessen (Lanis) to popular calendar apps. It generates RFC 5545 compliant .ics files.

## Features

- **Universal Export**: Compatible with Google Calendar, Apple Calendar, and Outlook.
- **Timezone Aware**: Correctly handles Europe/Berlin time and Daylight Saving Time (Sommerzeit/Winterzeit).
- **Smart Breaks**: Optionally calculates and inserts breaks (Pausen) based on gaps between lessons.
- **Holiday Integration**: Fetches public holidays and school vacations to keep the calendar clean.
- **Encoding Safe**: Forces UTF-8 with BOM to ensure German Umlaute render correctly in Excel and Outlook.

## Installation

1. Install a userscript manager like [Tampermonkey](https://www.tampermonkey.net/) or [Violentmonkey](https://violentmonkey.github.io/).
2. [Click here to install the script](https://raw.githubusercontent.com/0iy/schulportal-hessen-timetable-export/main/schulportal-exporter.user.js).
3. Confirm the installation.

## Usage

1. Log in to the Schulportal.
2. Open your timetable ("Mein Stundenplan").
3. Click the new Export button (next to "Drucken").
4. Select your classes and export format.

## Troubleshooting

### Google Calendar Import Failed?
Use the web interface (Settings > Import) rather than the mobile app for the first import.

### Weird characters in Outlook?
The script applies a Byte Order Mark (BOM) automatically. If issues persist, try importing to Google Calendar first.

## License

MIT License.

## Disclaimer
Not affiliated with Schulportal Hessen.
