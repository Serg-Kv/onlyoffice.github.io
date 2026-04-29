# Change Log

## 1.0.4

* Bridge `executeMethod` to `Asc.Editor.callMethod` in web hosts where the legacy plugin API is not exposed.
* Guard store rating rendering when a plugin has no `rating` payload.

## 1.0.3

* Add a web fallback for modal-dependent flows when `PluginWindow` is unavailable.
* Prevent web builds from crashing on info/settings/report windows and allow direct apply with saved settings.

## 1.0.2

* Add modal panel config for packaged internal windows.
* Fix child window initialization in desktop packaged plugin builds.

## 1.0.1

* Fix modal panel runtime loading in packaged plugin builds.
* Bundle local plugin UI runtime for internal windows.

## 1.0.0

* Initial release.
