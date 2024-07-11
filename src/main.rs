//! List the special "known folders" on a Windows system, and their locations.
//!
//! See [Known Folders](https://learn.microsoft.com/en-us/windows/win32/shell/known-folders).

use core::ffi::c_void;
use std::collections::HashMap;
use std::string::FromUtf16Error;

use thiserror::Error;

use windows::core::{Error as WindowsError, GUID, PWSTR};
use windows::Win32::System::Com::{
    CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, CLSCTX_INPROC_SERVER,
    COINIT_APARTMENTTHREADED,
};
use windows::Win32::UI::Shell::{
    IKnownFolder, IKnownFolderManager, KnownFolderManager, KF_FLAG_ALIAS_ONLY, KF_FLAG_CREATE,
    KF_FLAG_DEFAULT, KF_FLAG_DEFAULT_PATH, KF_FLAG_DONT_UNEXPAND, KF_FLAG_DONT_VERIFY,
    KF_FLAG_FORCE_APPCONTAINER_REDIRECTION, KF_FLAG_FORCE_APP_DATA_REDIRECTION,
    KF_FLAG_FORCE_PACKAGE_REDIRECTION, KF_FLAG_INIT, KF_FLAG_NOT_PARENT_RELATIVE, KF_FLAG_NO_ALIAS,
    KF_FLAG_NO_PACKAGE_REDIRECTION, KF_FLAG_RETURN_FILTER_REDIRECTION_TARGET,
    KF_FLAG_SIMPLE_IDLIST, KNOWNFOLDER_DEFINITION, KNOWN_FOLDER_FLAG,
};

#[derive(Debug, Error)]
enum FlagParseError {
    #[error("No options are recognized (got {0:?})")]
    UnrecognizedOption(String),
    #[error("Unrecognized flag name: {0}")]
    UnrecognizedFlag(String),
    #[error("Refusing to attempt to pass {0} for ALL known folders (dangerous)")]
    BannedFlag(String),
}

/// Makes an array of pairs of each name as a string with the resolved name.
macro_rules! named {
    ($($ident:ident),* $(,)?) => {
        [$(
            (stringify!($ident), $ident),
        )*]
    };
}

/// Pairs of known folder flags' symbolic names and the flag values.
const NAMED_KF_FLAGS: &[(&str, KNOWN_FOLDER_FLAG)] = &named!(
    KF_FLAG_DEFAULT,
    KF_FLAG_FORCE_APP_DATA_REDIRECTION,
    KF_FLAG_RETURN_FILTER_REDIRECTION_TARGET,
    KF_FLAG_FORCE_PACKAGE_REDIRECTION,
    KF_FLAG_NO_PACKAGE_REDIRECTION,
    KF_FLAG_FORCE_APPCONTAINER_REDIRECTION,
    KF_FLAG_CREATE, // Though we will refuse to attempt it.
    KF_FLAG_DONT_VERIFY,
    KF_FLAG_DONT_UNEXPAND,
    KF_FLAG_NO_ALIAS,
    KF_FLAG_INIT, // Though we will refuse, as it is only meaningful with KF_FLAG_CREATE.
    KF_FLAG_DEFAULT_PATH,
    KF_FLAG_NOT_PARENT_RELATIVE,
    KF_FLAG_SIMPLE_IDLIST,
    KF_FLAG_ALIAS_ONLY,
);

/// Flags we refuse to pass, because we would be passing them for ALL known folders.
const BANNED_KF_FLAGS: &[KNOWN_FOLDER_FLAG] = &[KF_FLAG_CREATE, KF_FLAG_INIT];

/// Convert an informal representation of a `KNOWN_FOLDER_FLAG` to the real name.
fn normalize_flag_name(flag_arg: &str) -> String {
    const PREFIX: &str = "KF_FLAG_";
    let upcased = flag_arg.to_uppercase();
    if upcased.starts_with(PREFIX) {
        upcased
    } else {
        format!("{PREFIX}{upcased}")
    }
}

/// Parse command line arguments as `KNOWN_FOLDER_FLAG` values.
///
/// Note that these represent how the operation of looking up a known folder's path
/// is customized. They do not identify specific known folders. (This program always
/// displays information about all registered known folders.)
///
/// This refuses to accept flags that would attempt to create directories for all
/// registered known folders that do not yet have them, or that would only be
/// meaningful in the presence of other flags that do this, since using this
/// diagnostic utility to create a potentially large number of directories is very
/// unlikely to be intended. To just see what the paths *would* all be if they were
/// created, the `KF_FLAG_DONT_VERIFY` flag can be used.
fn read_args_as_kf_flags() -> Result<KNOWN_FOLDER_FLAG, FlagParseError> {
    let table: HashMap<_, _> = HashMap::from_iter(NAMED_KF_FLAGS.iter().cloned());
    let mut flags = KF_FLAG_DEFAULT;
    assert!(flags.0 == 0, "Bug: Default flags are somehow nonzero!");

    for flag_arg in std::env::args().skip(1) {
        if flag_arg.starts_with('-') {
            return Err(FlagParseError::UnrecognizedOption(flag_arg));
        }

        let flag_name = normalize_flag_name(&flag_arg);
        match table.get(flag_name.as_str()) {
            None => return Err(FlagParseError::UnrecognizedFlag(flag_name)),
            Some(flag) if BANNED_KF_FLAGS.contains(flag) => {
                return Err(FlagParseError::BannedFlag(flag_name));
            }
            Some(flag) => flags |= *flag,
        }
    }

    for banned_flag in BANNED_KF_FLAGS {
        assert!(
            !flags.contains(*banned_flag),
            "Bug: Other flags somehow combined to form banned flag {banned_flag:?}"
        );
    }

    Ok(flags)
}

/// Guard type that initializes COM on the current thread and uninitializes it on drop.
struct ComInit;

impl ComInit {
    fn new() -> Result<Self, WindowsError> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }.ok()?;
        Ok(Self)
    }
}

impl Drop for ComInit {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

/// Free a `PWSTR` with `CoTaskMemFree`.
fn co_free_pwstr(pwstr: PWSTR) {
    unsafe { CoTaskMemFree(Some(pwstr.as_ptr().cast::<c_void>())) };
}

/// Owner of a `PWSTR` that must be freed with `CoTaskMemFree`.
struct CoStr {
    pwstr: PWSTR,
}

impl CoStr {
    fn new(pwstr: PWSTR) -> Self {
        Self { pwstr }
    }

    fn to_string(&self) -> Result<String, FromUtf16Error> {
        unsafe { self.pwstr.to_string() }
    }
}

impl Drop for CoStr {
    fn drop(&mut self) {
        co_free_pwstr(self.pwstr);
    }
}

/// Owner of `IKnownFolderManager::GetFolderIds` results.
///
/// On drop, this calls `CoTaskMemFree` on the block of GUIDs representing known folders.
struct KnownFolderIds {
    pkfid: *mut GUID,
    count: u32,
}

impl KnownFolderIds {
    fn new(kf_manager: &IKnownFolderManager) -> Result<Self, WindowsError> {
        let mut pkfid = std::ptr::null_mut();
        let mut count = 0;
        unsafe { kf_manager.GetFolderIds(&mut pkfid, &mut count)? };
        Ok(Self { pkfid, count })
    }

    fn as_slice(&self) -> &[GUID] {
        unsafe { std::slice::from_raw_parts(self.pkfid, self.count as usize) }
    }
}

impl Drop for KnownFolderIds {
    fn drop(&mut self) {
        unsafe { CoTaskMemFree(Some(self.pkfid.cast::<c_void>())) };
    }
}

/// Owner of a `KNOWNFOLDER_DEFINITION` that frees its dynamic strings on drop.
struct KnownFolderDefinition {
    fields: KNOWNFOLDER_DEFINITION,
}

impl KnownFolderDefinition {
    fn of(folder: &IKnownFolder) -> Result<Self, WindowsError> {
        let mut fields = KNOWNFOLDER_DEFINITION::default();
        unsafe { folder.GetFolderDefinition(&mut fields)? };
        Ok(Self { fields })
    }
}

impl Drop for KnownFolderDefinition {
    fn drop(&mut self) {
        // The windows crate does not provide FreeKnownFolderDefinitionFields, possibly
        // due to it being an __inline function. This frees each of the fields that is a
        // pointer to a string, which is equivalent to FreeKnownFolderDefinitionFields.
        co_free_pwstr(self.fields.pszName);
        co_free_pwstr(self.fields.pszDescription);
        co_free_pwstr(self.fields.pszRelativePath);
        co_free_pwstr(self.fields.pszParsingName);
        co_free_pwstr(self.fields.pszTooltip);
        co_free_pwstr(self.fields.pszLocalizedName);
        co_free_pwstr(self.fields.pszIcon);
        co_free_pwstr(self.fields.pszSecurity);
    }
}

/// A known folder name and either its retrieved path or an error.
struct NamedPath {
    name: String,
    try_path: Result<String, WindowsError>,
}

/// Get all known folder names and either paths or an error from getting the path.
fn get_named_paths(flags: KNOWN_FOLDER_FLAG) -> Result<Vec<NamedPath>, WindowsError> {
    let mut named_paths = vec![];
    unsafe {
        let kf_manager: IKnownFolderManager =
            CoCreateInstance(&KnownFolderManager, None, CLSCTX_INPROC_SERVER)?;
        for id in KnownFolderIds::new(&kf_manager)?.as_slice() {
            let folder = kf_manager.GetFolder(id)?;
            let name = KnownFolderDefinition::of(&folder)?
                .fields
                .pszName
                .to_string()?;
            let try_path = match folder.GetPath(flags.0 as u32) {
                Ok(pwstr) => Ok(CoStr::new(pwstr).to_string()?),
                Err(e) => Err(e),
            };
            named_paths.push(NamedPath { name, try_path });
        }
    }
    Ok(named_paths)
}

/// Displays a table of each known folder name with its path or why the path is unavailable.
fn print_table(named_paths: Vec<NamedPath>) {
    let name_width_estimate = named_paths
        .iter()
        .map(|np| np.name.chars().count())
        .max()
        .unwrap_or(0);

    for NamedPath { name, try_path } in named_paths {
        let path_item = try_path.unwrap_or_else(|e| format!("[{}]", e.message()));
        println!("{name:<name_width_estimate$}  {path_item}");
    }
}

/// Use the `IKnownFolder` API to retrieve information, and print it in tabular form.
fn run(flags: KNOWN_FOLDER_FLAG) -> Result<(), WindowsError> {
    let mut named_paths = get_named_paths(flags)?;
    named_paths.sort_by(|a, b| a.name.cmp(&b.name));
    print_table(named_paths);
    Ok(())
}

fn main() -> Result<(), WindowsError> {
    // Parse arguments and bail out if we cannot proceed.
    let flags = read_args_as_kf_flags().unwrap_or_else(|e| {
        eprintln!("Error: {e}");
        std::process::exit(2);
    });

    // To use `IKnownFolder`, we must have COM initialize on this thread.
    let _com = ComInit::new()?;

    // Use those flags to access the COM API for known folders and list them out.
    run(flags)
}
