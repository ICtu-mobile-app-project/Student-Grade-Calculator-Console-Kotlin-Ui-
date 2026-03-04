# first-kot

Minimal Android Kotlin starter template for coding and debugging on a physical Android device.

## Prerequisites

- Android SDK + platform tools (`adb`) in PATH
- JDK 17
- Gradle Wrapper files (generated in the next step)

## One-time setup

1. Connect device and enable Developer Options + USB debugging.
2. Trust the computer on the device.
3. Verify connection:

   ```powershell
   adb devices
   ```

## Generate Gradle wrapper

From project root:

```powershell
gradle wrapper
```

This creates `gradlew`, `gradlew.bat`, and `gradle/wrapper/*`.

## Build + install

Use VS Code task:

- `Android: installDebug`

Or command line:

```powershell
.\gradlew.bat installDebug
```

## Debug on external device (VS Code)

1. Run task `Android: start debug app`
2. Start launch config `Android: Attach JVM (port 8700)`

This attaches the Java debugger to the app process via JDWP forwarding.
