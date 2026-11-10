# forge-m365-sp-profiles

SharePoint user profile operations for the forge-m365 SDK, hand-ported from
PnP.js [`packages/sp/profiles`](https://github.com/pnp/pnpjs/tree/8ee2375d/packages/sp/profiles).

## Port source

- **Ref:** PnP.js commit `8ee2375d3ce720926774cc87168e1251c9ed584`
- **Path:** `packages/sp/profiles/`

## Covered operations

### PeopleManager (`_api/sp.userprofiles.peoplemanager`)

| Operation | HTTP | REST path |
|-----------|------|-----------|
| `get_edit_profile_link` | GET | `/EditProfileLink` |
| `get_is_my_people_list_public` | GET | `/IsMyPeopleListPublic` |
| `am_i_followed_by(login)` | GET | `/amifollowedby(@v)` |
| `am_i_following(login)` | GET | `/amifollowing(@v)` |
| `get_followed_tags(max)` | GET | `/getfollowedtags(max)` |
| `get_followers_for(login)` | GET | `/getfollowersfor(@v)` |
| `get_my_followers` | GET | `/getmyfollowers` |
| `get_my_properties` | GET | `/getmyproperties` |
| `get_people_followed_by(login)` | GET | `/getpeoplefollowedby(@v)` |
| `get_properties_for(login)` | GET | `/getpropertiesfor(@v)` |
| `get_trending_tags` | GET | `/gettrendingtags` |
| `get_user_profile_property_for(login, prop)` | GET | `/getuserprofilepropertyfor(...)` |
| `is_following(follower, followee)` | GET | `/isfollowing(...)` |
| `hide_suggestion(login)` | POST | `/hidesuggestion(@v)` |
| `set_single_value_profile_property(...)` | POST | `/SetSingleValueProfileProperty` |
| `set_multi_valued_profile_property(...)` | POST | `/SetMultiValuedProfileProperty` |
| `set_my_profile_pic(bytes)` | POST | `/setmyprofilepicture` |

### ProfileLoader (`_api/sp.userprofiles.profileloader`)

| Operation | HTTP | REST path |
|-----------|------|-----------|
| `get_user_profile` | POST | `.../getuserprofile` |
| `get_owner_user_profile` | POST | `.../getowneruserprofile` |
| `create_personal_site(interactive)` | POST | `.../createpersonalsiteenque(...)` |
| `create_personal_site_enqueue_bulk(emails)` | POST | `.../createpersonalsiteenqueuebulk` |
| `share_all_social_data(share)` | POST | `.../shareallsocialdata(...)` |

### ClientPeoplePicker

| Operation | HTTP | REST path |
|-----------|------|-----------|
| `client_people_picker_resolve_user(params)` | POST | `.../clientpeoplepickerresolveuser` |
| `client_people_picker_search_user(params)` | POST | `.../clientpeoplepickersearchuser` |
