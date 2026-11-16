(function(){const a=document.createElement("link").relList;if(a&&a.supports&&a.supports("modulepreload"))return;for(const o of document.querySelectorAll('link[rel="modulepreload"]'))i(o);new MutationObserver(o=>{for(const l of o)if(l.type==="childList")for(const s of l.addedNodes)s.tagName==="LINK"&&s.rel==="modulepreload"&&i(s)}).observe(document,{childList:!0,subtree:!0});function n(o){const l={};return o.integrity&&(l.integrity=o.integrity),o.referrerPolicy&&(l.referrerPolicy=o.referrerPolicy),o.crossOrigin==="use-credentials"?l.credentials="include":o.crossOrigin==="anonymous"?l.credentials="omit":l.credentials="same-origin",l}function i(o){if(o.ep)return;o.ep=!0;const l=n(o);fetch(o.href,l)}})();const oe="modulepreload",re=function(t,a){return new URL(t,a).href},q={},H=function(a,n,i){let o=Promise.resolve();if(n&&n.length>0){const s=document.getElementsByTagName("link"),r=document.querySelector("meta[property=csp-nonce]"),u=r?.nonce||r?.getAttribute("nonce");o=Promise.allSettled(n.map(m=>{if(m=re(m,i),m in q)return;q[m]=!0;const h=m.endsWith(".css"),O=h?'[rel="stylesheet"]':"";if(!!i)for(let M=s.length-1;M>=0;M--){const L=s[M];if(L.href===m&&(!h||L.rel==="stylesheet"))return}else if(document.querySelector(`link[href="${m}"]${O}`))return;const y=document.createElement("link");if(y.rel=h?"stylesheet":oe,h||(y.as="script"),y.crossOrigin="",y.href=m,u&&y.setAttribute("nonce",u),document.head.appendChild(y),h)return new Promise((M,L)=>{y.addEventListener("load",M),y.addEventListener("error",()=>L(new Error(`Unable to preload CSS for ${m}`)))})}))}function l(s){const r=new Event("vite:preloadError",{cancelable:!0});if(r.payload=s,window.dispatchEvent(r),!r.defaultPrevented)throw s}return o.then(s=>{for(const r of s||[])r.status==="rejected"&&l(r.reason);return a().catch(l)})},S=window.location.hostname==="localhost"||window.location.hostname==="127.0.0.1"?"":"https://localhost:3978",$="teams_outlook_sid";function ie(){try{const t=localStorage.getItem($);if(t)return t;const a=sessionStorage.getItem($);if(a)return localStorage.setItem($,a),a}catch{}return null}class W extends Error{constructor(a="Local server at localhost:3978 is unreachable"){super(a),this.name="ServerUnreachableError"}}let P=null,F=ie();function le(t){P=t}function z(t){F=t;try{t?(localStorage.setItem($,t),sessionStorage.setItem($,t)):(localStorage.removeItem($),sessionStorage.removeItem($))}catch{}}function de(t,a){const n=new Headers(t);return a&&!n.has("Content-Type")&&n.set("Content-Type","application/json"),P?(n.set("Authorization",`Bearer ${P}`),n.set("x-token-type","teams-sso")):F&&n.set("x-session-id",F),n}async function Z(t,a={}){const n=de(a.headers,!!a.body);let i;try{i=await fetch(`${S}${t}`,{...a,headers:n,credentials:"include"})}catch(o){throw o instanceof DOMException&&o.name==="AbortError"?o:new W}if(!i.ok){let o=`Request failed (${i.status})`;try{const l=await i.json();l.error?.message&&(o=l.error.message)}catch{}throw new Error(o)}return i}async function f(t,a={}){return await(await Z(t,a)).json()}function _(){return f("/api/auth/status")}function ce(){return f("/api/auth/logout",{method:"POST"})}function Q(){return f("/api/mail/folders")}function k(t,a){const n=new URLSearchParams({folder:t.folder,unreadOnly:String(t.unreadOnly),top:String(t.top??25),skip:String(t.skip??0)});return t.search.trim()&&n.set("search",t.search.trim()),f(`/api/mail/messages?${n.toString()}`,{signal:a})}function ue(t,a){return f(`/api/mail/messages/${encodeURIComponent(t)}`,{signal:a})}function Y(t,a){return f(`/api/mail/messages/${encodeURIComponent(t)}/read`,{method:"PATCH",body:JSON.stringify({isRead:a})})}async function me(t,a,n){const o=await(await Z(`/api/mail/messages/${encodeURIComponent(t)}/attachments/${encodeURIComponent(a)}`)).blob(),l=URL.createObjectURL(o),s=document.createElement("a");s.href=l,s.download=n||"attachment",document.body.appendChild(s),s.click(),s.remove(),setTimeout(()=>URL.revokeObjectURL(l),1e3)}function pe(t){return f(`/api/mail/messages/${encodeURIComponent(t)}`,{method:"DELETE"})}function fe(t){return f(`/api/mail/messages/${encodeURIComponent(t)}/archive`,{method:"POST"})}function he(t,a,n){return f(`/api/mail/messages/${encodeURIComponent(t)}/forward`,{method:"POST",body:JSON.stringify({to:a,comment:n})})}function ge(t,a,n=!1,i){return f(`/api/mail/messages/${encodeURIComponent(t)}/reply`,{method:"POST",body:JSON.stringify({comment:a,replyAll:n,attachments:i})})}function be(t){return f("/api/mail/messages/send",{method:"POST",body:JSON.stringify(t)})}function ye(t,a){return f(`/api/mail/recipients?q=${encodeURIComponent(t)}`,{signal:a})}function ve(t=!1){return f(`/api/app/version${t?"?check=1":""}`)}function we(){return f("/api/app/update",{method:"POST"})}function V(t){return t==="dark"||t==="contrast"?t:"default"}function R(t,a,n){return Promise.race([t,new Promise((i,o)=>setTimeout(()=>o(new Error(`${n} timeout`)),a))])}async function $e(t){if(!(window.parent!==window||!!window.opener))return{inTeams:!1,theme:"default",ssoToken:null};try{const{app:n,authentication:i}=await H(async()=>{const{app:r,authentication:u}=await import("./src-CZ2ouLw7.js");return{app:r,authentication:u}},[],import.meta.url);await R(n.initialize(),500,"Teams init");const o=await R(n.getContext(),1e3,"Teams getContext"),l=V(o.app?.theme);n.registerOnThemeChangeHandler(r=>{t(V(r))});let s=null;try{s=await R(i.getAuthToken(),1200,"Teams getAuthToken")}catch{s=null}return{inTeams:!0,theme:l,ssoToken:s}}catch{return{inTeams:!1,theme:"default",ssoToken:null}}}const X=[{id:"inbox",label:"Inbox"},{id:"sentitems",label:"Sent"},{id:"drafts",label:"Drafts"},{id:"archive",label:"Archive"},{id:"deleteditems",label:"Deleted"}],v=document.getElementById("root"),E=new Map;let e={theme:"default",inTeams:!1,auth:null,serverUnreachable:!1,folder:"inbox",folders:[],unreadOnly:!1,search:"",messages:[],hasMore:!1,loadingMore:!1,selectedId:null,selectedDetail:null,loadingList:!1,loadingDetail:!1,replyOpen:!1,replyAll:!1,sendingReply:!1,replyAttachments:[],forwardOpen:!1,sendingForward:!1,composeOpen:!1,sendingCompose:!1,composeAttachments:[],updatingApp:!1,versionInfo:null,error:null},D=null,C=null,B=null,x=null,U=null,g=[],b=-1;function d(t){return t?t.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}function ee(t){const a=new Date(t);return Number.isNaN(a.getTime())?t:a.toLocaleString([],{month:"short",day:"numeric",hour:"2-digit",minute:"2-digit"})}function K(t){e.theme=t,document.documentElement.setAttribute("data-theme",t)}function Ee(t){return t<1024?`${t} B`:t<1024*1024?`${(t/1024).toFixed(1)} KB`:`${(t/(1024*1024)).toFixed(1)} MB`}async function G(t){if(!t||t.length===0)return[];const a=[];for(const n of Array.from(t).slice(0,10)){const i=await n.arrayBuffer(),o=new Uint8Array(i);let l="";const s=32768;for(let r=0;r<o.length;r+=s)l+=String.fromCharCode(...o.subarray(r,r+s));a.push({name:n.name,contentType:n.type||"application/octet-stream",contentBytes:btoa(l)})}return a}function T(t,a){return t.length===0?"":`
    <div class="attachment-chips">
      ${t.map((n,i)=>`
          <span class="attachment-chip">
            <span>${d(n.name)}</span>
            <button type="button" class="chip-remove" data-remove-att="${a}" data-att-idx="${i}" aria-label="Remove ${d(n.name)}">&times;</button>
          </span>`).join("")}
    </div>`}function I(){if(e.messages.length===0)return`<div class="empty-state">${e.loadingList?"Loading messages…":"No messages found."}</div>`;const t=e.messages.map(n=>{const i=n.from?.emailAddress.name||n.from?.emailAddress.address||"Unknown",o=n.isRead?"":"unread",l=e.selectedId===n.id?"selected":"",s=n.isRead?"":'<span class="unread-dot" aria-label="Unread"></span>',r=n.hasAttachments?'<span class="att-badge" title="Has attachments">📎</span>':"";return`
        <button class="message-item ${o} ${l}" data-id="${d(n.id)}">
          <div class="message-item-top">
            <span class="sender-row">
              ${s}
              <span class="sender-name">${d(i)}</span>
            </span>
            <span class="message-time">${r}${d(ee(n.receivedDateTime))}</span>
          </div>
          <div class="message-subject">${d(n.subject||"(No subject)")}</div>
          <div class="message-preview">${d(n.bodyPreview)}</div>
        </button>`}).join(""),a=e.hasMore?`<div class="load-more-wrap"><button id="load-more-btn" class="btn load-more-btn" ${e.loadingMore?"disabled":""}>${e.loadingMore?"Loading older…":"Load older messages"}</button></div>`:"";return t+a}function te(){if(e.loadingDetail)return'<div class="empty-state">Loading message…</div>';const t=e.selectedDetail;if(!t)return'<div class="empty-state">Select a message to read or reply.</div>';const a=t.from?.emailAddress.name?`${t.from.emailAddress.name} <${t.from.emailAddress.address}>`:t.from?.emailAddress.address??"",n=t.body.contentType.toLowerCase()==="html"?t.body.content:`<pre class="email-body-text">${d(t.body.content)}</pre>`,i=t.attachments&&t.attachments.length>0?`
      <div class="detail-attachments">
        ${t.attachments.map(s=>`
            <button type="button" class="attachment-chip attachment-download" data-download-att="${d(s.id)}" data-att-name="${d(s.name)}">
              <span>📎 ${d(s.name)}</span>
              <span class="att-size">${d(Ee(s.size))}</span>
            </button>`).join("")}
      </div>`:"",o=e.replyOpen?`
      <div class="reply-box">
        <strong>${e.replyAll?"Reply All":"Reply"}</strong>
        <textarea id="reply-text" class="textarea" placeholder="Write your reply… (Cmd/Ctrl+Enter to send)"></textarea>
        <div class="attachment-picker-row">
          <label class="btn btn-sm file-picker-label">
            Attach files
            <input id="reply-files" type="file" multiple hidden />
          </label>
          <div id="reply-att-pills">${T(e.replyAttachments,"reply")}</div>
        </div>
        <div class="detail-actions">
          <button id="send-reply-btn" class="btn btn-primary" ${e.sendingReply?"disabled":""}>
            ${e.sendingReply?"Sending…":"Send Reply"}
          </button>
          <button id="cancel-reply-btn" class="btn">Cancel</button>
        </div>
      </div>`:"",l=e.forwardOpen?`
      <div class="reply-box">
        <strong>Forward</strong>
        <input id="forward-to" type="text" class="field-input" placeholder="To (comma-separated emails)" autocomplete="off" />
        <textarea id="forward-comment" class="textarea" placeholder="Add a note (optional)… (Cmd/Ctrl+Enter to send)"></textarea>
        <div class="detail-actions">
          <button id="send-forward-btn" class="btn btn-primary" ${e.sendingForward?"disabled":""}>
            ${e.sendingForward?"Forwarding…":"Send Forward"}
          </button>
          <button id="cancel-forward-btn" class="btn">Cancel</button>
        </div>
      </div>`:"";return`
    <div class="detail-card">
      <div class="detail-header">
        <div>
          <h2 class="detail-subject">${d(t.subject||"(No subject)")}</h2>
          <div class="detail-meta">
            <span><strong>From:</strong> ${d(a)}</span>
            <span><strong>Date:</strong> ${d(ee(t.receivedDateTime))}</span>
          </div>
        </div>
        <div class="detail-actions">
          <button id="toggle-read-btn" class="btn">${t.isRead?"Mark unread":"Mark read"}</button>
          <button id="reply-btn" class="btn btn-primary">Reply</button>
          <button id="reply-all-btn" class="btn">Reply All</button>
          <button id="forward-btn" class="btn">Forward</button>
          <button id="archive-btn" class="btn">Archive</button>
          <button id="delete-btn" class="btn btn-danger">Delete</button>
        </div>
      </div>
      ${i}
      <div class="email-body">${n}</div>
      ${o}
      ${l}
    </div>`}function c(){if(e.serverUnreachable){v.innerHTML=`
      <div class="empty-state">
        <h2>Local Server Unavailable</h2>
        <p>Cannot reach the Outlook plugin backend at <code>${d(S||window.location.origin)}</code>. Make sure the local service is running.</p>
        <button id="retry-server-btn" class="btn btn-primary">Retry connection</button>
      </div>`;return}const{auth:t}=e;if(!t){v.innerHTML='<div class="empty-state">Loading Outlook Mail…</div>';return}if(!t.authenticated){const h=t.entraConfigured?e.inTeams?'<button id="teams-login-btn" class="btn btn-primary">Sign in with Microsoft</button>':`<a class="btn btn-primary" href="${S}/api/auth/login">Sign in with Microsoft</a>`:"<p>Configure <code>AZURE_CLIENT_ID</code> and <code>AZURE_CLIENT_SECRET</code> in <code>.env</code>, or enable <code>DEV_MOCK_GRAPH=true</code>.</p>";v.innerHTML=`
      <div class="empty-state">
        <h2>Outlook Mail for Teams</h2>
        <p>Sign in with your Microsoft 365 account to view and manage your Outlook mailbox.</p>
        ${h}
      </div>`;return}const a=t.user?`<span class="user-label">${d(t.user.displayName)} (${d(t.user.mail)})</span>`:"",n=t.mode==="session"?'<button id="logout-btn" class="btn">Sign out</button>':"",i=e.error?`<div class="error-banner" role="alert"><span>${d(e.error)}</span><button id="dismiss-err" class="btn">Dismiss</button></div>`:"",o=X.map(h=>{const O=e.folders.find(y=>y.id===h.id),j=O&&O.unreadItemCount>0?` <span class="folder-badge">${O.unreadItemCount}</span>`:"";return`
      <button role="tab" aria-selected="${e.folder===h.id}"
        class="folder-tab ${e.folder===h.id?"active":""}"
        data-folder="${h.id}">${h.label}${j}</button>`}).join(""),l=e.composeOpen?`
      <div class="modal-backdrop" role="dialog" aria-modal="true" aria-label="New Mail">
        <form id="compose-form" class="modal">
          <h3>New Mail</h3>
          <div class="autocomplete-wrap">
            <input id="compose-to" type="text" class="field-input" placeholder="To (comma-separated emails)" autocomplete="off" required />
            <div id="compose-suggestions" class="autocomplete-list" hidden></div>
          </div>
          <input id="compose-subject" type="text" class="field-input" placeholder="Subject" required />
          <textarea id="compose-body" class="textarea" placeholder="Message body… (Cmd/Ctrl+Enter to send)"></textarea>
          <div class="attachment-picker-row">
            <label class="btn btn-sm file-picker-label">
              Attach files
              <input id="compose-files" type="file" multiple hidden />
            </label>
            <div id="compose-att-pills">${T(e.composeAttachments,"compose")}</div>
          </div>
          <div class="modal-footer">
            <button type="button" id="cancel-compose-btn" class="btn">Cancel</button>
            <button type="submit" class="btn btn-primary" ${e.sendingCompose?"disabled":""}>
              ${e.sendingCompose?"Sending…":"Send"}
            </button>
          </div>
        </form>
      </div>`:"",s=e.versionInfo?`v${d(e.versionInfo.version)}`:"v1.0.4",r=e.versionInfo?`v${d(e.versionInfo.version)} (${d(e.versionInfo.commit)})`:"v1.0.4",u=e.versionInfo?.updateAvailable?`${e.versionInfo.behindCount} update(s) available`:"Up to date",m=e.versionInfo?.updateAvailable?'<span class="update-dot" title="Update available"></span>':"";v.innerHTML=`
    <header class="topbar">
      <div class="topbar-left">
        <h1 class="app-title">Outlook Mail</h1>
        <span class="mode-badge">${e.inTeams?"Teams Tab":"Browser"} · ${t.mode==="mock"?"Mock Graph":"Graph API"}</span>
        <div class="app-brand" tabindex="0">
          <span class="mode-badge">${s} ${m}</span>
          <div class="version-popover">
            <div class="version-card">
              <div><strong>Outlook Mail</strong> ${r}</div>
              <div class="version-meta">${d(u)}</div>
              <button id="app-update-btn" class="btn btn-primary" ${e.updatingApp?"disabled":""}>
                ${e.updatingApp?"Updating…":"Update"}
              </button>
            </div>
          </div>
        </div>
      </div>
      <div class="topbar-actions">
        ${a}
        <button id="new-mail-btn" class="btn btn-primary">New Mail</button>
        <button id="refresh-btn" class="btn" ${e.loadingList?"disabled":""}>
          ${e.loadingList?"Refreshing…":"Refresh"}
        </button>
        <a class="btn" href="https://outlook.office.com/mail/" target="_blank" rel="noopener noreferrer">Open OWA ↗</a>
        ${n}
      </div>
    </header>
    ${i}
    <div class="toolbar">
      <div class="folder-tabs" role="tablist" aria-label="Mail folders">${o}</div>
      <div class="search-controls">
        <label class="checkbox-label">
          <input id="unread-toggle" type="checkbox" ${e.unreadOnly?"checked":""} />
          Unread only
        </label>
        <input id="search-input" type="search" class="search-input" placeholder="Search mail…"
          aria-label="Search mail" value="${d(e.search)}" />
      </div>
    </div>
    <main class="workspace">
      <section id="message-list" class="message-list" aria-label="Messages">${I()}</section>
      <section id="detail-pane" class="detail-pane" aria-label="Message detail">${te()}</section>
    </main>
    ${l}`}function p(){const t=document.getElementById("message-list"),a=document.getElementById("detail-pane");if(!t||!a){c();return}t.innerHTML=I(),a.innerHTML=te();const n=document.getElementById("refresh-btn");n&&(n.disabled=e.loadingList,n.textContent=e.loadingList?"Refreshing…":"Refresh")}function Ae(){for(const t of X){const a=document.querySelector(`.folder-tab[data-folder="${t.id}"]`);if(!a)continue;const n=e.folders.find(o=>o.id===t.id),i=n&&n.unreadItemCount>0?` <span class="folder-badge">${n.unreadItemCount}</span>`:"";a.innerHTML=`${t.label}${i}`}}async function ae(){try{const t=await Q();e.folders=t.folders,Ae()}catch{}}async function w(){D?.abort(),D=new AbortController,e.loadingList=!0,e.hasMore=!1,e.error=null,p();try{const[t]=await Promise.all([k({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search},D.signal),ae()]);e.messages=t.messages,e.hasMore=t.hasMore}catch(t){if(t.name==="AbortError")return;e.error=t instanceof Error?t.message:"Failed to load messages",c();return}finally{e.loadingList=!1}p()}async function ne(){if(e.loadingMore||e.loadingList||!e.hasMore)return;e.loadingMore=!0;const t=document.getElementById("message-list");t&&(t.innerHTML=I());try{const n=await k({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search,skip:e.messages.length,top:25}),i=new Set(e.messages.map(l=>l.id)),o=n.messages.filter(l=>!i.has(l.id));e.messages=[...e.messages,...o],e.hasMore=n.hasMore&&o.length>0}catch(n){e.error=n instanceof Error?n.message:"Failed to load older messages"}finally{e.loadingMore=!1}const a=document.getElementById("message-list");a&&(a.innerHTML=I())}async function Te(t){C?.abort(),C=new AbortController,e.selectedId=t,e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,e.error=null;const a=E.get(t);if(a)e.selectedDetail=a,e.loadingDetail=!1,p();else{e.loadingDetail=!0,p();try{const n=await ue(t,C.signal);E.set(t,n),e.selectedDetail=n}catch(n){if(n.name==="AbortError")return;e.error=n instanceof Error?n.message:"Failed to open message",e.loadingDetail=!1,c();return}finally{e.loadingDetail=!1}}e.selectedDetail&&!e.selectedDetail.isRead?(e.selectedDetail.isRead=!0,e.messages=e.messages.map(n=>n.id===t?{...n,isRead:!0}:n),p(),Y(t,!0).catch(()=>{})):p()}v.addEventListener("click",async t=>{const a=t.target,n=a.closest(".message-item[data-id]");if(n?.dataset.id){Te(n.dataset.id);return}const i=a.closest(".folder-tab[data-folder]");if(i?.dataset.folder){e.folder=i.dataset.folder,e.selectedId=null,e.selectedDetail=null,c(),w();return}if(a.id==="retry-server-btn"){e.serverUnreachable=!1,e.auth=null,c(),N();return}const o=a.closest("[data-download-att]");if(o?.dataset.downloadAtt&&e.selectedDetail){try{await me(e.selectedDetail.id,o.dataset.downloadAtt,o.dataset.attName??"attachment")}catch(s){e.error=s instanceof Error?s.message:"Failed to download attachment",c()}return}const l=a.closest("[data-remove-att]");if(l?.dataset.removeAtt){const s=Number(l.dataset.attIdx);if(l.dataset.removeAtt==="compose"){e.composeAttachments.splice(s,1);const r=document.getElementById("compose-att-pills");r&&(r.innerHTML=T(e.composeAttachments,"compose"))}else if(l.dataset.removeAtt==="reply"){e.replyAttachments.splice(s,1);const r=document.getElementById("reply-att-pills");r&&(r.innerHTML=T(e.replyAttachments,"reply"))}return}if(a.id==="load-more-btn"){ne();return}if(a.id==="refresh-btn")E.clear(),w();else if(a.id==="app-update-btn"){e.updatingApp=!0,c();try{await we();for(let s=0;s<20;s++)if(await new Promise(u=>setTimeout(u,500)),await fetch(`${S}/api/health`).then(u=>u.ok).catch(()=>!1)){window.location.reload();return}window.location.reload()}catch(s){e.updatingApp=!1,e.error=s instanceof Error?s.message:"Update failed",c()}}else if(a.closest(".autocomplete-item[data-email]")){const s=a.closest(".autocomplete-item[data-email]");se(s.dataset.email)}else if(a.id==="new-mail-btn")e.composeOpen=!0,c();else if(a.id==="cancel-compose-btn")e.composeOpen=!1,c();else if(a.id==="dismiss-err")e.error=null,c();else if(a.id==="teams-login-btn")try{const{authentication:s}=await H(async()=>{const{authentication:m}=await import("./src-CZ2ouLw7.js");return{authentication:m}},[],import.meta.url),r=await s.authenticate({url:`${S||window.location.origin}/api/auth/login?teams=1`,width:600,height:560});r&&z(r);const u=await _();e.auth=u,c(),u.authenticated&&w()}catch(s){e.error=s instanceof Error?s.message:"Microsoft sign-in failed",c()}else if(a.id==="logout-btn")await ce(),z(null),E.clear(),e.auth=await _(),e.messages=[],e.selectedDetail=null,c();else if(a.id==="toggle-read-btn"&&e.selectedDetail){const s=e.selectedDetail,r=!s.isRead;s.isRead=r,e.messages=e.messages.map(u=>u.id===s.id?{...u,isRead:r}:u),p();try{await Y(s.id,r)}catch(u){e.error=u instanceof Error?u.message:"Failed to update read status",c()}}else if(a.id==="reply-btn")e.replyAll=!1,e.replyOpen=!0,e.forwardOpen=!1,p();else if(a.id==="reply-all-btn")e.replyAll=!0,e.replyOpen=!0,e.forwardOpen=!1,p();else if(a.id==="cancel-reply-btn")e.replyOpen=!1,e.replyAttachments=[],p();else if(a.id==="send-reply-btn"&&e.selectedDetail){const r=document.getElementById("reply-text")?.value.trim()??"";if(!r)return;e.sendingReply=!0,p();try{await ge(e.selectedDetail.id,r,e.replyAll,e.replyAttachments.length>0?e.replyAttachments:void 0),e.replyOpen=!1,e.replyAttachments=[],e.sendingReply=!1,await w()}catch(u){e.sendingReply=!1,e.error=u instanceof Error?u.message:"Failed to send reply",c()}}else if(a.id==="forward-btn")e.forwardOpen=!0,e.replyOpen=!1,p();else if(a.id==="cancel-forward-btn")e.forwardOpen=!1,p();else if(a.id==="send-forward-btn"&&e.selectedDetail){const s=document.getElementById("forward-to")?.value??"",r=document.getElementById("forward-comment")?.value.trim()??"",u=s.split(",").map(m=>m.trim()).filter(Boolean);if(u.length===0){e.error="Enter at least one recipient email to forward",c();return}e.sendingForward=!0,p();try{await he(e.selectedDetail.id,u,r),e.forwardOpen=!1,e.sendingForward=!1,await w()}catch(m){e.sendingForward=!1,e.error=m instanceof Error?m.message:"Failed to forward message",c()}}else if(a.id==="archive-btn"&&e.selectedDetail){const s=e.selectedDetail.id;try{await fe(s),E.delete(s),e.messages=e.messages.filter(r=>r.id!==s),e.selectedId=null,e.selectedDetail=null,p()}catch(r){e.error=r instanceof Error?r.message:"Failed to archive message",c()}}else if(a.id==="delete-btn"&&e.selectedDetail){const s=e.selectedDetail.id;try{await pe(s),E.delete(s),e.messages=e.messages.filter(r=>r.id!==s),e.selectedId=null,e.selectedDetail=null,p()}catch(r){e.error=r instanceof Error?r.message:"Failed to delete message",c()}}});v.addEventListener("change",async t=>{const a=t.target;if(a.id==="unread-toggle")e.unreadOnly=a.checked,w();else if(a.id==="compose-files"){const n=await G(a.files);e.composeAttachments=[...e.composeAttachments,...n].slice(0,10);const i=document.getElementById("compose-att-pills");i&&(i.innerHTML=T(e.composeAttachments,"compose"))}else if(a.id==="reply-files"){const n=await G(a.files);e.replyAttachments=[...e.replyAttachments,...n].slice(0,10);const i=document.getElementById("reply-att-pills");i&&(i.innerHTML=T(e.replyAttachments,"reply"))}});function A(){const t=document.getElementById("compose-suggestions");if(t){if(g.length===0){t.hidden=!0,t.innerHTML="";return}t.hidden=!1,t.innerHTML=g.map((a,n)=>`
      <button type="button" class="autocomplete-item ${n===b?"active":""}" data-email="${d(a.address)}">
        <span><strong>${d(a.name)}</strong> &lt;${d(a.address)}&gt;</span>
        <span class="autocomplete-sub">${d(a.source)}</span>
      </button>`).join("")}}function se(t){const a=document.getElementById("compose-to");if(!a)return;const n=a.value.split(",");n[n.length-1]=` ${t}`,a.value=`${n.map(i=>i.trim()).filter(Boolean).join(", ")}, `,g=[],b=-1,A(),a.focus()}v.addEventListener("input",t=>{const a=t.target;if(a.id==="search-input")e.search=a.value,x&&window.clearTimeout(x),x=window.setTimeout(()=>void w(),200);else if(a.id==="compose-to"){const n=a.value.split(",").pop()?.trim()??"";if(U&&window.clearTimeout(U),n.length<1){g=[],b=-1,A();return}U=window.setTimeout(async()=>{B?.abort(),B=new AbortController;try{g=(await ye(n,B.signal)).suggestions,b=g.length>0?0:-1,A()}catch{}},150)}});v.addEventListener("keydown",t=>{t.target.id!=="compose-to"||g.length===0||(t.key==="ArrowDown"?(t.preventDefault(),b=(b+1)%g.length,A()):t.key==="ArrowUp"?(t.preventDefault(),b=(b-1+g.length)%g.length,A()):t.key==="Enter"&&b>=0?(t.preventDefault(),se(g[b].address)):t.key==="Escape"&&(g=[],b=-1,A()))});v.addEventListener("submit",async t=>{const a=t.target;if(a.id!=="compose-form")return;t.preventDefault();const n=document.getElementById("compose-to")?.value??"",i=document.getElementById("compose-subject")?.value.trim()??"",o=document.getElementById("compose-body")?.value??"",l=n.split(",").map(r=>r.trim()).filter(Boolean);if(l.length===0||!i)return;e.sendingCompose=!0;const s=a.querySelector('button[type="submit"]');s&&(s.disabled=!0,s.textContent="Sending…");try{await be({to:l,subject:i,body:o,attachments:e.composeAttachments.length>0?e.composeAttachments:void 0}),e.composeOpen=!1,e.composeAttachments=[],e.sendingCompose=!1,g=[],b=-1,c(),w()}catch(r){e.sendingCompose=!1,s&&(s.disabled=!1,s.textContent="Send"),e.error=r instanceof Error?r.message:"Failed to send email",c()}});async function J(){if(!(!e.auth?.authenticated||document.hidden||e.composeOpen||e.replyOpen||e.loadingList||e.loadingMore))try{const[t]=await Promise.all([k({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search}),ae()]);if(e.messages.length<=25)e.messages=t.messages,e.hasMore=t.hasMore;else{const n=new Set(t.messages.map(o=>o.id)),i=e.messages.filter(o=>!n.has(o.id));e.messages=[...t.messages,...i]}const a=document.getElementById("message-list");if(a){const n=a.scrollTop;a.innerHTML=I(),a.scrollTop=n}}catch{}}const Oe=3e4;v.addEventListener("scroll",t=>{const a=t.target;a?.id!=="message-list"||!e.hasMore||e.loadingMore||e.loadingList||a.scrollTop+a.clientHeight>=a.scrollHeight-120&&ne()},!0);async function N(){if(window.location.pathname==="/auth-done"){const a=new URLSearchParams(window.location.search).get("sid")??"";try{const{app:n,authentication:i}=await H(async()=>{const{app:o,authentication:l}=await import("./src-CZ2ouLw7.js");return{app:o,authentication:l}},[],import.meta.url);await n.initialize(),i.notifySuccess(a)}catch{window.close()}return}const t=await $e(K);e.inTeams=t.inTeams,K(t.theme),le(t.ssoToken);try{const a=_(),n=k({folder:e.folder,unreadOnly:e.unreadOnly,search:e.search}).catch(()=>null),i=Q().catch(()=>null),[o,l,s]=await Promise.all([a,n,i]);e.serverUnreachable=!1,e.auth=o,s&&(e.folders=s.folders),o.authenticated&&l?(e.messages=l.messages,e.hasMore=l.hasMore):o.authenticated&&!l&&w(),c()}catch(a){if(a instanceof W){e.serverUnreachable=!0,c(),window.setTimeout(()=>{e.serverUnreachable&&N()},4e3);return}e.error=a instanceof Error?a.message:"Failed to initialize",c()}ve(!0).then(a=>{e.versionInfo=a,c()}).catch(()=>{}),window.setInterval(()=>void J(),Oe),document.addEventListener("visibilitychange",()=>{document.hidden||J()}),document.addEventListener("keydown",a=>{(a.metaKey||a.ctrlKey)&&a.key==="Enter"?e.composeOpen?(a.preventDefault(),document.getElementById("compose-form")?.requestSubmit()):e.replyOpen?(a.preventDefault(),document.getElementById("send-reply-btn")?.click()):e.forwardOpen&&(a.preventDefault(),document.getElementById("send-forward-btn")?.click()):a.key==="Escape"&&(e.composeOpen?(e.composeOpen=!1,e.composeAttachments=[],c()):(e.replyOpen||e.forwardOpen)&&(e.replyOpen=!1,e.replyAttachments=[],e.forwardOpen=!1,p()))})}N();
