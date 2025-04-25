import os
import zipfile


def create_plugin_archive(name_plugin: str):
    if os.path.exists(f"{name_plugin}.plugin"):
        os.remove(f"{name_plugin}.plugin")

    files_to_archive = []

    for root, dirs, files in os.walk("."):

        for file in files:
            if file != "main.py":
                files_to_archive.append(os.path.join(root, file))

    print(files_to_archive)

    with zipfile.ZipFile(f"{name_plugin}.zip", "w") as zipf:
        for file in files_to_archive:
            zipf.write(file)

    os.rename(f"{name_plugin}.zip", f"{name_plugin}.plugin")


create_plugin_archive("R7_SEARCH")
