import { createClient } from '@supabase/supabase-js'

const projectUrl = import.meta.env.VITE_SUPABASE_URL
const publishableKey = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY

// Some mobile browsers filter custom cross-origin headers such as `apikey` and
// `authorization` before the request reaches Supabase. The publishable key is
// intentionally public, so also send it as a query parameter and keep the REST
// request headers to the small set needed by PostgREST.
const mobileCompatibleFetch: typeof fetch = async (input, init) => {
  const request = new Request(input, init)
  const url = new URL(request.url)
  const isSupabaseRest = url.hostname.endsWith('.supabase.co') && url.pathname.startsWith('/rest/')
  if (!isSupabaseRest || !publishableKey) return fetch(request)

  url.searchParams.set('apikey', publishableKey)
  const headers = new Headers(request.headers)
  headers.delete('apikey')
  headers.delete('authorization')
  headers.delete('x-client-info')
  headers.delete('content-profile')

  const body = request.method === 'GET' || request.method === 'HEAD' ? undefined : await request.clone().arrayBuffer()
  return fetch(url.toString(), {
    method: request.method,
    headers,
    body,
    credentials: request.credentials,
    cache: request.cache,
    redirect: request.redirect,
    referrer: request.referrer,
    referrerPolicy: request.referrerPolicy,
    mode: request.mode,
    signal: request.signal,
  })
}

export const supabase: any = projectUrl && publishableKey
  ? createClient(projectUrl, publishableKey, { global: { fetch: mobileCompatibleFetch } })
  : null
