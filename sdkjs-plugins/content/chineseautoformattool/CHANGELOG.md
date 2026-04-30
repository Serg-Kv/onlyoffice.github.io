# Change Log

## 1.0.7

* Point panel-relative translation loading back to the root `translations` folder via `base href="../"`.
* Keep the remote panel runtime from `1.0.6` unchanged and avoid duplicating locale files under `panels/`.

## 1.0.6

* Load panel runtime files from the external `sdkjs-plugins/v1` endpoint instead of bundled local `scripts/V1`.
* Keep `panels/config.json` unchanged while testing remote runtime behavior for child windows.

## 1.0.5

* Restore the pre-web-fallback implementation after the modal window fix.
* Keep the packaged modal window behavior from the `1.0.2` generation and drop the later web-only experiments.

## 1.0.2

* Add modal panel config for packaged internal windows.
* Fix child window initialization in desktop packaged plugin builds.

## 1.0.1

* Fix modal panel runtime loading in packaged plugin builds.
* Bundle local plugin UI runtime for internal windows.

## 1.0.0

* Initial release.
