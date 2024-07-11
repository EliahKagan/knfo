//! List the special "known folders" on a Windows system, and their locations.
//!
//! See [Known Folders](https://learn.microsoft.com/en-us/windows/win32/shell/known-folders).

use core::ffi::c_void;
use std::string::FromUtf16Error;

use windows::{
    core::{Error, GUID, PWSTR},
    Win32::{
        System::Com::{
            CoCreateInstance, CoInitializeEx, CoTaskMemFree, CoUninitialize, CLSCTX_INPROC_SERVER,
            COINIT_APARTMENTTHREADED,
        },
        UI::Shell::{
            IKnownFolder, IKnownFolderManager, KnownFolderManager, KF_FLAG_DEFAULT,
            KNOWNFOLDER_DEFINITION, KNOWN_FOLDER_FLAG,
        },
    },
};

struct ComInit;

impl ComInit {
    fn new() -> Result<Self, Error> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED) }.ok()?;
        Ok(Self)
    }
}

impl Drop for ComInit {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

fn co_free_pwstr(pwstr: PWSTR) {
    unsafe { CoTaskMemFree(Some(pwstr.as_ptr().cast::<c_void>())) };
}

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

struct KnownFolderIds {
    pkfid: *mut GUID,
    count: u32,
}

impl KnownFolderIds {
    fn new(kf_manager: &IKnownFolderManager) -> Result<Self, Error> {
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

struct KnownFolderDefinition {
    fields: KNOWNFOLDER_DEFINITION,
}

impl KnownFolderDefinition {
    fn of(folder: &IKnownFolder) -> Result<Self, Error> {
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

struct NamedPath {
    name: String,
    try_path: Result<String, Error>,
}

fn get_named_paths(flags: KNOWN_FOLDER_FLAG) -> Result<Vec<NamedPath>, Error> {
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

fn main() -> Result<(), Error> {
    let _com = ComInit::new()?;

    // TODO: Rather than always use KF_FLAG_DEFAULT, accept flags as command-line arguments.
    let mut named_paths = get_named_paths(KF_FLAG_DEFAULT)?;
    named_paths.sort_by(|a, b| a.name.cmp(&b.name));
    print_table(named_paths);

    Ok(())
}
