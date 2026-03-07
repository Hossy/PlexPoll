# PlexPoll

HomeSeer HS4 VB.NET script that polls Plex `/status/sessions` and updates mapped HomeSeer media devices/features.

## Purpose

This script is a patch/workaround for (hopefully temporary) broken functionality in BLPlex.

- It is intended to work with existing HomeSeer devices/features already created by BLPlex.
- It does not create new HomeSeer devices.
- It updates BLPlex device state/feature data by polling Plex sessions directly.

Related forum thread:

- https://forums.homeseer.com/forum/media-plug-ins/media-discussion/blplex-blade/1745813-initialized-but-no-devices

## License

GNU GPL v3. See `LICENSE`.

Copyright (C) John Greg Hossbach

## Files

- `PlexSessionPoll.vb`: main HomeSeer script.
- `PlexSessionPoll.ini.example`: example config file (copy to HS4 `Config\PlexSessionPoll.ini`).

## Install

1. Copy `PlexSessionPoll.vb` to:
   - `C:\Program Files (x86)\HomeSeer HS4\scripts\`
2. Copy `PlexSessionPoll.ini.example` to:
   - `C:\Program Files (x86)\HomeSeer HS4\Config\PlexSessionPoll.ini`
3. Edit `PlexSessionPoll.ini` values:
   - `[Plex] Token`
   - `[Plex] Server` (hostname or IP)
   - `[Plex] Port`
   - `[Players]` mappings (`machineIdentifier=ParentDeviceRef`)
4. Create a recurring HomeSeer event to run `PlexSessionPoll.vb`.

## Config Notes

`PlexSessionPoll.ini` supports:

- `UseHttps` (`true/false`)
- `showLastPlayedWhenStopped` (`true/false`)
- Timeout values:
  - `TimeoutMs`
  - `ResolveTimeoutMs`
  - `ConnectTimeoutMs`
  - `SendTimeoutMs`
  - `ReceiveTimeoutMs`

Fallback behavior (if not set in `PlexSessionPoll.ini`):

- `Server` falls back to `hspi_BLPlex.ini` -> `[Settings] plexIpAddress`
- `Port` falls back to `hspi_BLPlex.ini` -> `[Settings] plexPort`
- `showLastPlayedWhenStopped` falls back to `hspi_BLPlex.ini` -> `[Settings] showLastPlayedWhenStopped`

## HomeSeer Scripting Reminder

This script is VB.NET syntax in HomeSeer:

- Do not use legacy `Set`/`Let` assignment.
- Method calls with arguments must use parentheses.

## Status Behavior

Parent device playback state is set by numeric value map (VSPairs), not by status strings.
