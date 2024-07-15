# knfo - Get info on Windows known folders

This is a Rust program that uses the [`IKnownFolderManager`](https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-iknownfoldermanager) COM interface to list the special [known folders](https://learn.microsoft.com/en-us/windows/win32/shell/known-folders) on a Windows system, and their locations. These are special locations like Program Files, a user's Documents directory, and so forth.

It uses the Windows API through the [`windows`](https://crates.io/crates/windows) crate.

## Usage

This always lists all [known folders](https://learn.microsoft.com/en-us/windows/win32/shell/known-folders) registered with the system, including those that are registered but do not currently exist, and including those that are not inherent to Windows but have been added by the user or a third-party application.

They are listed alphabetized by their names for readability, even though this is not likely to be the order the system returns them in. Note that these are their names in the known folders system, and should not be confused with their paths (when present), or with the symbolic constants that exist for some of them.

When errors occur in obtaining information about a known folder, whether due to a location not existing on disk or for any other reason, the error is reported `[in brackets]` in place of a path.

Command-line arguments, if passed, are taken to be custom [`KNOWN_FOLDER_FLAG`](https://learn.microsoft.com/en-us/windows/win32/api/shlobj_core/ne-shlobj_core-known_folder_flag) values. These can be passed with or without the leading text `KF_FLAG_`. Pass one flag per argument. Passing none is equivalent to `KF_FLAG_DEFAULT`.

This program will refuse to proceed if `KF_FLAG_CREATE` is one of the flags, because this is a diagnostic tool, and as such it is unlikely that creating (or attempting to create) every possibly currently registered known folder is wanted.

## License

[0BSD](LICENSE)
