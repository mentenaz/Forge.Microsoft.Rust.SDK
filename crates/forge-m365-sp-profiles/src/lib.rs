pub mod types;

use forge_m365_core::{registered_operations, Client, OperationEntry, Result, Surface};
use forge_m365_macros::pnp_operation;
use serde_json::Value;

// PeopleManager operations
#[pnp_operation(id = "sp.profiles.get_edit_profile_link", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_edit_profile_link_op() {}

#[pnp_operation(id = "sp.profiles.get_is_my_people_list_public", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_is_my_people_list_public_op() {}

#[pnp_operation(id = "sp.profiles.am_i_followed_by", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn am_i_followed_by_op() {}

#[pnp_operation(id = "sp.profiles.am_i_following", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn am_i_following_op() {}

#[pnp_operation(id = "sp.profiles.get_followed_tags", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_followed_tags_op() {}

#[pnp_operation(id = "sp.profiles.get_followers_for", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_followers_for_op() {}

#[pnp_operation(id = "sp.profiles.get_my_followers", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_my_followers_op() {}

#[pnp_operation(id = "sp.profiles.get_my_properties", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_my_properties_op() {}

#[pnp_operation(id = "sp.profiles.get_people_followed_by", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_people_followed_by_op() {}

#[pnp_operation(id = "sp.profiles.get_properties_for", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_properties_for_op() {}

#[pnp_operation(id = "sp.profiles.get_trending_tags", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_trending_tags_op() {}

#[pnp_operation(id = "sp.profiles.get_user_profile_property_for", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_user_profile_property_for_op() {}

#[pnp_operation(id = "sp.profiles.hide_suggestion", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn hide_suggestion_op() {}

#[pnp_operation(id = "sp.profiles.is_following", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn is_following_op() {}

#[pnp_operation(id = "sp.profiles.set_single_value_profile_property", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn set_single_value_profile_property_op() {}

#[pnp_operation(id = "sp.profiles.set_multi_valued_profile_property", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn set_multi_valued_profile_property_op() {}

#[pnp_operation(id = "sp.profiles.set_my_profile_pic", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn set_my_profile_pic_op() {}

// ProfileLoader operations
#[pnp_operation(id = "sp.profiles.get_user_profile", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_user_profile_op() {}

#[pnp_operation(id = "sp.profiles.get_owner_user_profile", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn get_owner_user_profile_op() {}

#[pnp_operation(id = "sp.profiles.create_personal_site", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn create_personal_site_op() {}

#[pnp_operation(id = "sp.profiles.create_personal_site_enqueue_bulk", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn create_personal_site_enqueue_bulk_op() {}

#[pnp_operation(id = "sp.profiles.share_all_social_data", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn share_all_social_data_op() {}

// ClientPeoplePicker operations
#[pnp_operation(id = "sp.profiles.client_people_picker_resolve_user", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn client_people_picker_resolve_user_op() {}

#[pnp_operation(id = "sp.profiles.client_people_picker_search_user", primary = Surface::SpRest, fallback = [])]
#[allow(dead_code)]
async fn client_people_picker_search_user_op() {}

fn entry(id: &'static str) -> &'static OperationEntry {
    registered_operations()
        .find(|e| e.id == id)
        .unwrap_or_else(|| {
            panic!("operation '{id}' not registered; is forge-m365-sp-profiles linked in?")
        })
}

fn profiles_url(site_url: &str) -> String {
    format!(
        "{}/_api/sp.userprofiles.peoplemanager",
        site_url.trim_end_matches('/')
    )
}

fn encode_login_name(name: &str) -> String {
    name.replace('\'', "''")
}

// ---------------------------------------------------------------------------
// PeopleManager — read operations
// ---------------------------------------------------------------------------

/// Gets the URL of the edit profile page for the current user.
///
/// Ported from PnPjs `Profiles.getEditProfileLink`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_edit_profile_link(client: &Client<'_>, site_url: &str) -> Result<String> {
    let url = format!("{}/EditProfileLink", profiles_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_edit_profile_link"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// A boolean that indicates whether the current user's "People I'm
/// Following" list is public.
///
/// Ported from PnPjs `Profiles.getIsMyPeopleListPublic`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_is_my_people_list_public(client: &Client<'_>, site_url: &str) -> Result<bool> {
    let url = format!("{}/IsMyPeopleListPublic", profiles_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_is_my_people_list_public"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// A boolean that indicates whether the current user is being followed
/// by the specified user.
///
/// Ported from PnPjs `Profiles.amIFollowedBy`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn am_i_followed_by(
    client: &Client<'_>,
    site_url: &str,
    login_name: &str,
) -> Result<bool> {
    let url = format!(
        "{}/amifollowedby(@v)?@v='{}'",
        profiles_url(site_url),
        encode_login_name(login_name),
    );
    let bytes = client
        .run_ladder(
            entry("sp.profiles.am_i_followed_by"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// A boolean that indicates whether the current user is following the
/// specified user.
///
/// Ported from PnPjs `Profiles.amIFollowing`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn am_i_following(client: &Client<'_>, site_url: &str, login_name: &str) -> Result<bool> {
    let url = format!(
        "{}/amifollowing(@v)?@v='{}'",
        profiles_url(site_url),
        encode_login_name(login_name),
    );
    let bytes = client
        .run_ladder(entry("sp.profiles.am_i_following"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets tags that the current user is following.
///
/// Ported from PnPjs `Profiles.getFollowedTags`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_followed_tags(
    client: &Client<'_>,
    site_url: &str,
    max_count: i64,
) -> Result<Vec<String>> {
    let url = format!("{}/getfollowedtags({max_count})", profiles_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_followed_tags"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the people who are following the specified user.
///
/// Ported from PnPjs `Profiles.getFollowersFor`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_followers_for(
    client: &Client<'_>,
    site_url: &str,
    login_name: &str,
) -> Result<Vec<Value>> {
    let url = format!(
        "{}/getfollowersfor(@v)?@v='{}'",
        profiles_url(site_url),
        encode_login_name(login_name),
    );
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_followers_for"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the people who are following the current user.
///
/// Ported from PnPjs `Profiles.myFollowers`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_my_followers(client: &Client<'_>, site_url: &str) -> Result<Vec<Value>> {
    let url = format!("{}/getmyfollowers", profiles_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_my_followers"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets user properties for the current user.
///
/// Ported from PnPjs `Profiles.myProperties`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_my_properties(client: &Client<'_>, site_url: &str) -> Result<Value> {
    let url = format!("{}/getmyproperties", profiles_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_my_properties"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the people who the specified user is following.
///
/// Ported from PnPjs `Profiles.getPeopleFollowedBy`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_people_followed_by(
    client: &Client<'_>,
    site_url: &str,
    login_name: &str,
) -> Result<Vec<Value>> {
    let url = format!(
        "{}/getpeoplefollowedby(@v)?@v='{}'",
        profiles_url(site_url),
        encode_login_name(login_name),
    );
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_people_followed_by"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets user properties for the specified user.
///
/// Ported from PnPjs `Profiles.getPropertiesFor`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_properties_for(
    client: &Client<'_>,
    site_url: &str,
    login_name: &str,
) -> Result<Value> {
    let url = format!(
        "{}/getpropertiesfor(@v)?@v='{}'",
        profiles_url(site_url),
        encode_login_name(login_name),
    );
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_properties_for"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the 20 most popular hash tags over the past week.
///
/// Ported from PnPjs `Profiles.trendingTags`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_trending_tags(
    client: &Client<'_>,
    site_url: &str,
) -> Result<types::HashTagCollection> {
    let url = format!("{}/gettrendingtags", profiles_url(site_url));
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_trending_tags"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the specified user profile property for the specified user.
///
/// Ported from PnPjs `Profiles.getUserProfilePropertyFor`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_user_profile_property_for(
    client: &Client<'_>,
    site_url: &str,
    login_name: &str,
    property_name: &str,
) -> Result<String> {
    let url = format!(
        "{}/getuserprofilepropertyfor(accountname=@v, propertyname='{}')?@v='{}'",
        profiles_url(site_url),
        property_name,
        encode_login_name(login_name),
    );
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_user_profile_property_for"),
            "GET",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// A boolean that indicates whether the first user is following the
/// second user.
///
/// Ported from PnPjs `Profiles.isFollowing`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn is_following(
    client: &Client<'_>,
    site_url: &str,
    follower: &str,
    followee: &str,
) -> Result<bool> {
    let url = format!(
        "{}/isfollowing(possiblefolloweraccountname=@v, possiblefolloweeaccountname=@y)?@v='{}'&@y='{}'",
        profiles_url(site_url),
        encode_login_name(follower),
        encode_login_name(followee),
    );
    let bytes = client
        .run_ladder(entry("sp.profiles.is_following"), "GET", &url, &[], None)
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

// ---------------------------------------------------------------------------
// PeopleManager — write operations
// ---------------------------------------------------------------------------

/// Removes the specified user from the current user's list of suggested
/// people to follow.
///
/// Ported from PnPjs `Profiles.hideSuggestion`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn hide_suggestion(client: &Client<'_>, site_url: &str, login_name: &str) -> Result<()> {
    let url = format!(
        "{}/hidesuggestion(@v)?@v='{}'",
        profiles_url(site_url),
        encode_login_name(login_name),
    );
    client
        .run_ladder(
            entry("sp.profiles.hide_suggestion"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}

/// Sets a single-value user profile property.
///
/// Ported from PnPjs `Profiles.setSingleValueProfileProperty`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn set_single_value_profile_property(
    client: &Client<'_>,
    site_url: &str,
    account_name: &str,
    property_name: &str,
    property_value: &str,
) -> Result<()> {
    let url = format!("{}/SetSingleValueProfileProperty", profiles_url(site_url));
    let body = serde_json::json!({
        "accountName": account_name,
        "propertyName": property_name,
        "propertyValue": property_value,
    });
    let payload = serde_json::to_vec(&body)?;
    client
        .run_ladder(
            entry("sp.profiles.set_single_value_profile_property"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(())
}

/// Sets a multi-valued user profile property.
///
/// Ported from PnPjs `Profiles.setMultiValuedProfileProperty`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn set_multi_valued_profile_property(
    client: &Client<'_>,
    site_url: &str,
    account_name: &str,
    property_name: &str,
    property_values: &[String],
) -> Result<()> {
    let url = format!("{}/SetMultiValuedProfileProperty", profiles_url(site_url));
    let body = serde_json::json!({
        "accountName": account_name,
        "propertyName": property_name,
        "propertyValues": property_values,
    });
    let payload = serde_json::to_vec(&body)?;
    client
        .run_ladder(
            entry("sp.profiles.set_multi_valued_profile_property"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(())
}

/// Uploads and sets the user profile picture. Users can upload a picture
/// to their own profile only.
///
/// `image_bytes` must be BMP, JPEG, or PNG format, up to 4.76 MB.
///
/// Ported from PnPjs `Profiles.setMyProfilePic`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn set_my_profile_pic(
    client: &Client<'_>,
    site_url: &str,
    image_bytes: &[u8],
) -> Result<()> {
    let url = format!("{}/setmyprofilepicture", profiles_url(site_url));
    client
        .run_ladder(
            entry("sp.profiles.set_my_profile_pic"),
            "POST",
            &url,
            &[("Content-Type", "application/octet-stream")],
            Some(image_bytes),
        )
        .await?;
    Ok(())
}

// ---------------------------------------------------------------------------
// ProfileLoader operations
// ---------------------------------------------------------------------------

/// Gets the user profile of the current user via ProfileLoader.
///
/// Ported from PnPjs `ProfileLoader.userProfile`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_user_profile(client: &Client<'_>, site_url: &str) -> Result<types::UserProfile> {
    let url = format!(
        "{}/_api/sp.userprofiles.profileloader/getprofileloader/getuserprofile",
        site_url.trim_end_matches('/'),
    );
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_user_profile"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Gets the user profile of the site owner via ProfileLoader.
///
/// Ported from PnPjs `ProfileLoader.ownerUserProfile`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn get_owner_user_profile(
    client: &Client<'_>,
    site_url: &str,
) -> Result<types::UserProfile> {
    let url = format!(
        "{}/_api/sp.userprofiles.profileloader/getowneruserprofile",
        site_url.trim_end_matches('/'),
    );
    let bytes = client
        .run_ladder(
            entry("sp.profiles.get_owner_user_profile"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Enqueues creating a personal site for the current user.
///
/// Ported from PnPjs `ProfileLoader.createPersonalSite`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn create_personal_site(
    client: &Client<'_>,
    site_url: &str,
    interactive_request: bool,
) -> Result<()> {
    let url = format!(
        "{}/_api/sp.userprofiles.profileloader/getprofileloader/getuserprofile/createpersonalsiteenque({interactive_request})",
        site_url.trim_end_matches('/'),
    );
    client
        .run_ladder(
            entry("sp.profiles.create_personal_site"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}

/// Provisions one or more users' personal sites. (My Site administrator
/// on SharePoint Online only.)
///
/// Ported from PnPjs `ProfileLoader.createPersonalSiteEnqueueBulk`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn create_personal_site_enqueue_bulk(
    client: &Client<'_>,
    site_url: &str,
    emails: &[String],
) -> Result<()> {
    let url = format!(
        "{}/_api/sp.userprofiles.profileloader/getprofileloader/createpersonalsiteenqueuebulk",
        site_url.trim_end_matches('/'),
    );
    let body = serde_json::json!({ "emailIDs": emails });
    let payload = serde_json::to_vec(&body)?;
    client
        .run_ladder(
            entry("sp.profiles.create_personal_site_enqueue_bulk"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(())
}

/// Sets the privacy settings for the current user's profile.
///
/// Ported from PnPjs `ProfileLoader.shareAllSocialData`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn share_all_social_data(client: &Client<'_>, site_url: &str, share: bool) -> Result<()> {
    let url = format!(
        "{}/_api/sp.userprofiles.profileloader/getprofileloader/getuserprofile/shareallsocialdata({share})",
        site_url.trim_end_matches('/'),
    );
    client
        .run_ladder(
            entry("sp.profiles.share_all_social_data"),
            "POST",
            &url,
            &[],
            None,
        )
        .await?;
    Ok(())
}

// ---------------------------------------------------------------------------
// ClientPeoplePicker operations
// ---------------------------------------------------------------------------

/// Resolves a user or group using the Client People Picker.
///
/// Ported from PnPjs `ClientPeoplePickerQuery.clientPeoplePickerResolveUser`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn client_people_picker_resolve_user(
    client: &Client<'_>,
    site_url: &str,
    query_params: &types::PeoplePickerQueryParameters,
) -> Result<types::PeoplePickerEntity> {
    let url = format!(
        "{}/_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface/clientpeoplepickerresolveuser",
        site_url.trim_end_matches('/'),
    );
    let body = serde_json::json!({ "queryParams": query_params });
    let payload = serde_json::to_vec(&body)?;
    let bytes = client
        .run_ladder(
            entry("sp.profiles.client_people_picker_resolve_user"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}

/// Searches for users or groups using the Client People Picker.
///
/// Ported from PnPjs `ClientPeoplePickerQuery.clientPeoplePickerSearchUser`
/// (`packages/sp/profiles/types.ts`) @ pnpjs `8ee2375d`.
pub async fn client_people_picker_search_user(
    client: &Client<'_>,
    site_url: &str,
    query_params: &types::PeoplePickerQueryParameters,
) -> Result<Vec<types::PeoplePickerEntity>> {
    let url = format!(
        "{}/_api/sp.ui.applicationpages.clientpeoplepickerwebserviceinterface/clientpeoplepickersearchuser",
        site_url.trim_end_matches('/'),
    );
    let body = serde_json::json!({ "queryParams": query_params });
    let payload = serde_json::to_vec(&body)?;
    let bytes = client
        .run_ladder(
            entry("sp.profiles.client_people_picker_search_user"),
            "POST",
            &url,
            &[],
            Some(&payload),
        )
        .await?;
    Ok(serde_json::from_slice(&bytes)?)
}
