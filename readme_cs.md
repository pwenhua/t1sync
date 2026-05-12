# T1Sync C# — Visual Studio Setup

The `.cs` files in this folder are direct ports of the Python scripts. They use only the .NET base class library (no NuGet packages) and target **.NET 6 or newer**.

## Files

| File | Purpose | Entry-point class |
| --- | --- | --- |
| `t1_auth.cs` | OAuth2 client_credentials → access_token | `T1Sync.T1AuthApp` |
| `asset_get.cs` | GET single asset by `testID` | `T1Sync.AssetGetApp` |
| `parse_meta.cs` | Iterate `parse_meta`, extract `[level, field, dataType]` | `T1Sync.ParseMetaApp` |
| `T1Client.cs` | `T1Client` class: `GetTokenAsync`, `FetchAssetAsync`, `SaveAssetAsync` | *(library — no Main)* |
| `test_t1client.cs` | Smoke test for `T1Client` | `T1Sync.TestT1ClientApp` |

All classes live in the `T1Sync` namespace.

## 1. Create the project

In Visual Studio:

1. **File → New → Project…**
2. Pick **Console App** (C#), Next.
3. Name it `T1Sync`, choose a location, Next.
4. Framework: **.NET 8.0** (or any .NET 6+). Click **Create**.

This produces a project with a default `Program.cs` containing top-level statements.

## 2. Add the source files

1. **Delete the generated `Program.cs`** (or empty it). The provided files have their own `Main` methods, and a project may only have one entry point.
2. In Solution Explorer, right-click the project → **Add → Existing Item…**
3. Select all five `.cs` files from this folder and add them.
   - If you'd rather copy than link, use **Add as Link** off (the default). Linking is fine too if you want the project to track edits in this folder.

## 3. Add `config.json` and copy it on build

The code reads `config.json` from the working directory at runtime.

1. Right-click the project → **Add → Existing Item…** → select `config.json`.
2. Select `config.json` in Solution Explorer, open **Properties** (F4):
   - **Build Action**: `Content`
   - **Copy to Output Directory**: `Copy if newer`

This places `config.json` next to the built `.exe` so `File.OpenRead("config.json")` resolves.

## 4. Pick the entry point

Because every app file declares its own `Main`, you must tell the compiler which one to use. Edit the `.csproj` (right-click project → **Edit Project File**) and add `<StartupObject>` inside the existing `<PropertyGroup>`:

```xml
<PropertyGroup>
  <OutputType>Exe</OutputType>
  <TargetFramework>net8.0</TargetFramework>
  <Nullable>enable</Nullable>
  <ImplicitUsings>enable</ImplicitUsings>

  <StartupObject>T1Sync.TestT1ClientApp</StartupObject>
</PropertyGroup>
```

Swap the value to run a different program:

| To run | Set `<StartupObject>` to |
| --- | --- |
| `t1_auth.cs` | `T1Sync.T1AuthApp` |
| `asset_get.cs` | `T1Sync.AssetGetApp` |
| `parse_meta.cs` | `T1Sync.ParseMetaApp` |
| `test_t1client.cs` | `T1Sync.TestT1ClientApp` |

Visual Studio also exposes this via **Project → Properties → Application → Startup object** if you'd rather click than edit XML.

## 5. Build and run

- Press **F5** (Debug) or **Ctrl+F5** (Run without debugger).
- Or via terminal: `dotnet run` from the project folder.

## Notes

- **HTTPS verification is disabled** in `T1Client.cs` (`ServerCertificateCustomValidationCallback`) to match the Python `verify=False`. Remove that callback for hosts with trusted certificates.
- **`SaveAssetAsync` performs a real write** in `TestT1ClientApp`. Comment out the `AssetSaveTestAsync` call in `test_t1client.cs` if you only want to exercise the read path.
- **Working directory**: Visual Studio runs the app with cwd = `bin\Debug\net8.0\`. `config.json` lands there because of the Copy to Output Directory setting in step 3.
- **Excel Output (`SaveMetaToExcel`)**: To comply with the "no NuGet packages" rule by default, this method requires you to manually install the `ClosedXML` package (`dotnet add package ClosedXML`) and uncomment `#define USE_CLOSEDXML` at the top of `T1Client.cs`.

## Optional: split into multiple startup projects

If you want every entry point runnable without editing `<StartupObject>`, create separate console projects in the same solution (one per `App` class) and reference a shared class-library project that holds `T1Auth`, `AssetGet`, `ParseMeta`, and `T1Client`. For most development this is overkill — switching `<StartupObject>` is faster.
