pub use forge_m365_auth as auth;
pub use forge_m365_core::{Client, Error, Ladder, OperationEntry, Result, Surface};
pub use forge_m365_macros::pnp_operation;

pub mod sp {
    pub use forge_m365_sp_files as files;
    pub use forge_m365_sp_folders as folders;
    pub use forge_m365_sp_lists as lists;
    pub use forge_m365_sp_search as search;
    pub use forge_m365_sp_site_users as site_users;
    pub use forge_m365_sp_sites as sites;
    pub use forge_m365_sp_webs as webs;
}
