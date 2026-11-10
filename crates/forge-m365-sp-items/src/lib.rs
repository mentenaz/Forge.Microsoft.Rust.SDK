pub mod item;
pub mod item_properties;
pub mod types;
pub mod versions;

use forge_m365_core::Surface;
use forge_m365_macros::pnp_operation;

// Item operations
#[pnp_operation(id = "sp.items.get_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_id_op() {}

#[pnp_operation(id = "sp.items.get_by_string_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_by_string_id_op() {}

#[pnp_operation(id = "sp.items.recycle", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn recycle_op() {}

#[pnp_operation(id = "sp.items.delete_with_params", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_with_params_op() {}

#[pnp_operation(id = "sp.items.get_wopi_frame_url", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_wopi_frame_url_op() {}

#[pnp_operation(id = "sp.items.validate_update_list_item", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn validate_update_list_item_op() {}

#[pnp_operation(id = "sp.items.get_parent_infos", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_parent_infos_op() {}

// Item property operations
#[pnp_operation(id = "sp.items.get_effective_base_permissions", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_effective_base_permissions_op() {}

#[pnp_operation(id = "sp.items.get_effective_base_permissions_for_ui", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_effective_base_permissions_for_ui_op() {}

#[pnp_operation(id = "sp.items.get_field_values_as_html", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_field_values_as_html_op() {}

#[pnp_operation(id = "sp.items.get_field_values_as_text", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_field_values_as_text_op() {}

#[pnp_operation(id = "sp.items.get_field_values_for_edit", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_field_values_for_edit_op() {}

// Version operations
#[pnp_operation(id = "sp.items.get_versions", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_versions_op() {}

#[pnp_operation(id = "sp.items.get_version_by_id", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_version_by_id_op() {}

#[pnp_operation(id = "sp.items.delete_version", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn delete_version_op() {}
