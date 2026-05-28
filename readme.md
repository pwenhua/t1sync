# T1Sync C# — Visual Studio Setup

NuGet install ClosedXML

NuGet install Microsoft.Data.SqlClient

#

Done — argparse _main() removed, replaced with Trans.py:


if __name__ == "__main__":
    import fire
    fire.Fire(Trans.from_config())
Once you pip install fire, every method on the instance is directly callable:


python Trans.py csv2xlsx        c:/temp/thin.csv "Tree"
python Trans.py template2thin   c:/temp/template.csv c:/temp/thin.csv --asset_type_only
python Trans.py thin2import     c:/temp/thin.csv c:/temp/import.csv
python Trans.py save_meta_to_json c:/temp/template.csv c:/temp/meta.json "Tree/Street Tree"
python Trans.py --help lists all methods; python Trans.py csv2xlsx -- --help shows that specific method's args.

The Trans.from_config() factory runs once at startup, so nominated_fields is loaded a single time and shared across whichever method you invoke. The class-based API (used by test.py and Form1.cs) is untouched.

Heads-up: fire isn't installed in your current env yet — pip install fire to enable the CLI. The class API works without it.

