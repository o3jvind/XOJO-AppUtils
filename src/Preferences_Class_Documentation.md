
## Preferences Class Documentation

### Overview

The `Preferences` class provides a cross-platform, API 2-compliant mechanism to store and retrieve user preferences in a local SQLite database. It automatically manages database creation, type-safe access methods, and memory caching for performance.

### Features

- ✅ Cross-platform: macOS, Windows (not tested), Linux (not testet)
- ✅ XOJO API 2 compliant
- ✅ Automatic creation of platform-appropriate preferences path and SQLite file
- ✅ Caches all values in memory for fast access
- ✅ Strong typing for common XOJO types
- ✅ Stores data in a single `preferences` table with columns: `key`, `value`, `type`

### Supported Types

| Type         | Set Method           | Get Method             | Notes |
|--------------|----------------------|-------------------------|-------|
| String       | `SetString(key, val)`| `GetString(key, default)`| - |
| Boolean      | `SetBoolean(key, val)`| `GetBoolean(key, default)`| Stored as "1"/"0" |
| Integer 32   | `SetInteger32(key, val)`| `GetInteger32(key, default)`| |
| Integer 64   | `SetInteger64(key, val)`| `GetInteger64(key, default)`| |
| Double       | `SetDouble(key, val)`| `GetDouble(key, default)`| |
| Single       | `SetSingle(key, val)`| `GetSingle(key, default)`| |
| Color        | `SetColor(key, val)`| `GetColor(key, default)`| Stored as "&cRRGGBB" |
| DateTime     | `SetDateTime(key, val)`| `GetDateTime(key, default)`| ISO 8601 format |
| Picture      | `SetPicture(key, val)`| `GetPicture(key, default)`| Stored as Base64 |
| Array        | `SetArray(key, val)`| `GetArray(key, default)`| Serialized via JSON |
| Dictionary   | `SetDictionary(key, val)`| `GetDictionary(key, default)`| Serialized via JSON |

### Usage
Check **PreferencesTest.RunAllPreferencesTests**

### Notes

- Preferences are cached in memory. Changes made via SQL outside of the class won't be reflected until next launch.
- All keys are case-sensitive.
- Default values are returned if keys do not exist.
