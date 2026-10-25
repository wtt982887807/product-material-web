// Header-only isolation for static hosting. No document or response is cached.
self.addEventListener('install',event=>event.waitUntil(self.skipWaiting()));
self.addEventListener('activate',event=>event.waitUntil(self.clients.claim()));
self.addEventListener('fetch',event=>{
  const request=event.request;
  if(new URL(request.url).origin!==self.location.origin||!['GET','HEAD'].includes(request.method))return;
  if(request.cache==='only-if-cached'&&request.mode!=='same-origin')return;
  event.respondWith(fetch(request).then(response=>{
    // Preserve opaque responses and redirects to another origin untouched.
    if(response.status===0||response.url&&new URL(response.url).origin!==self.location.origin)return response;
    const headers=new Headers(response.headers);
    headers.set('Cross-Origin-Opener-Policy','same-origin');
    headers.set('Cross-Origin-Embedder-Policy','require-corp');
    return new Response(response.body,{status:response.status,statusText:response.statusText,headers});
  }));
});
