use proc_macro::TokenStream;
use quote::quote;
use syn::spanned::Spanned;
use syn::{
    parse::{Parse, ParseStream},
    parse_macro_input, Expr, ItemFn, Token,
};

struct LadderArgs {
    id: String,
    primary: Option<Expr>,
    fallback: Vec<Expr>,
}

mod kw {
    syn::custom_keyword!(id);
    syn::custom_keyword!(primary);
    syn::custom_keyword!(fallback);
}

impl Parse for LadderArgs {
    fn parse(input: ParseStream) -> syn::Result<Self> {
        let mut id = None;
        let mut primary = None;
        let mut fallback = Vec::new();

        while !input.is_empty() {
            let lookahead = input.lookahead1();
            if lookahead.peek(kw::id) {
                input.parse::<kw::id>()?;
                input.parse::<Token![=]>()?;
                let lit: syn::LitStr = input.parse()?;
                if id.is_some() {
                    return Err(syn::Error::new(lit.span(), "duplicate `id`"));
                }
                id = Some(lit.value());
            } else if lookahead.peek(kw::primary) {
                input.parse::<kw::primary>()?;
                input.parse::<Token![=]>()?;
                if primary.is_some() {
                    return Err(input.error("duplicate `primary`"));
                }
                primary = Some(input.parse::<Expr>()?);
            } else if lookahead.peek(kw::fallback) {
                input.parse::<kw::fallback>()?;
                input.parse::<Token![=]>()?;
                let content;
                syn::bracketed!(content in input);
                fallback = content
                    .parse_terminated(Expr::parse, Token![,])?
                    .into_iter()
                    .collect();
            } else {
                return Err(lookahead.error());
            }
            if !input.is_empty() {
                input.parse::<Token![,]>()?;
            }
        }

        Ok(LadderArgs {
            id: id.ok_or_else(|| input.error("missing required `id = \"...\"`"))?,
            primary,
            fallback,
        })
    }
}

#[proc_macro_attribute]
pub fn pnp_operation(attr: TokenStream, item: TokenStream) -> TokenStream {
    let args: LadderArgs = match syn::parse(attr) {
        Ok(a) => a,
        Err(e) => return e.to_compile_error().into(),
    };
    let func = parse_macro_input!(item as ItemFn);
    let fn_name = &func.sig.ident;

    let Some(primary) = args.primary else {
        return syn::Error::new(
            func.sig.span(),
            "#[pnp_operation] requires a `primary = ...` surface expression",
        )
        .to_compile_error()
        .into();
    };
    let fallback = &args.fallback;
    let id = &args.id;

    let expanded = quote! {
        #func

        forge_m365_core::inventory::submit! {
            forge_m365_core::OperationEntry {
                id: #id,
                operation_path: concat!(module_path!(), "::", stringify!(#fn_name)),
                ladder: forge_m365_core::Ladder {
                    primary: #primary,
                    fallback: &[#(#fallback),*],
                },
            }
        }
    };

    TokenStream::from(expanded)
}
