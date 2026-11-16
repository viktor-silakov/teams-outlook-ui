(function(){const s=document.createElement("link").relList;if(s&&s.supports&&s.supports("modulepreload"))return;for(const a of document.querySelectorAll('link[rel="modulepreload"]'))r(a);new MutationObserver(a=>{for(const i of a)if(i.type==="childList")for(const o of i.addedNodes)o.tagName==="LINK"&&o.rel==="modulepreload"&&r(o)}).observe(document,{childList:!0,subtree:!0});function n(a){const i={};return a.integrity&&(i.integrity=a.integrity),a.referrerPolicy&&(i.referrerPolicy=a.referrerPolicy),a.crossOrigin==="use-credentials"?i.credentials="include":a.crossOrigin==="anonymous"?i.credentials="omit":i.credentials="same-origin",i}function r(a){if(a.ep)return;a.ep=!0;const i=n(a);fetch(a.href,i)}})();const J="modulepreload",K=function(e,s){return new URL(e,s).href},B={},x=function(s,n,r){let a=Promise.resolve();if(n&&n.length>0){const o=document.getElementsByTagName("link"),l=document.querySelector("meta[property=csp-nonce]"),h=l?.nonce||l?.getAttribute("nonce");a=Promise.allSettled(n.map(u=>{if(u=K(u,r),u in B)return;B[u]=!0;const m=u.endsWith(".css"),G=m?'[rel="stylesheet"]':"";if(!!r)for(let T=o.length-1;T>=0;T--){const E=o[T];if(E.href===u&&(!m||E.rel==="stylesheet"))return}else if(document.querySelector(`link[href="${u}"]${G}`))return;const y=document.createElement("link");if(y.rel=m?"stylesheet":J,m||(y.as="script"),y.crossOrigin="",y.href=u,h&&y.setAttribute("nonce",h),document.head.appendChild(y),m)return new Promise((T,E)=>{y.addEventListener("load",T),y.addEventListener("error",()=>E(new Error(`Unable to preload CSS for ${u}`)))})}))}function i(o){const l=new Event("vite:preloadError",{cancelable:!0});if(l.payload=o,window.dispatchEvent(l),!l.defaultPrevented)throw o}return a.then(o=>{for(const l of o||[])l.status==="rejected"&&i(l.reason);return s().catch(i)})},S=window.location.hostname==="localhost"||window.location.hostname==="127.0.0.1"?"":"https://localhost:3978";let D=null,M=sessionStorage.getItem("teams_outlook_sid");function W(e){D=e}function N(e){M=e,e?sessionStorage.setItem("teams_outlook_sid",e):sessionStorage.removeItem("teams_outlook_sid")}async function b(e,s={}){const n=new Headers(s.headers);s.body&&!n.has("Content-Type")&&n.set("Content-Type","application/json"),D?(n.set("Authorization",`Bearer ${D}`),n.set("x-token-type","teams-sso")):M&&n.set("x-session-id",M);const r=await fetch(`${S}${e}`,{...s,headers:n,credentials:"include"});if(!r.ok){let a=`Request failed (${r.status})`;try{const i=await r.json();i.error?.message&&(a=i.error.message)}catch{}throw new Error(a)}return await r.json()}function C(){return b("/api/auth/status")}function Z(){return b("/api/auth/logout",{method:"POST"})}function _(e,s){const n=new URLSearchParams({folder:e.folder,unreadOnly:String(e.unreadOnly),top:String(e.top??25),skip:String(e.skip??0)});return e.search.trim()&&n.set("search",e.search.trim()),b(`/api/mail/messages?${n.toString()}`,{signal:s})}function Q(e,s){return b(`/api/mail/messages/${encodeURIComponent(e)}`,{signal:s})}function q(e,s){return b(`/api/mail/messages/${encodeURIComponent(e)}/read`,{method:"PATCH",body:JSON.stringify({isRead:s})})}function X(e,s,n=!1){return b(`/api/mail/messages/${encodeURIComponent(e)}/reply`,{method:"POST",body:JSON.stringify({comment:s,replyAll:n})})}function Y(e){return b("/api/mail/messages/send",{method:"POST",body:JSON.stringify(e)})}function ee(e,s){return b(`/api/mail/recipients?q=${encodeURIComponent(e)}`,{signal:s})}function te(e=!1){return b(`/api/app/version${e?"?check=1":""}`)}function se(){return b("/api/app/update",{method:"POST"})}function U(e){return e==="dark"||e==="contrast"?e:"default"}async function ne(e){if(!(window.parent!==window||!!window.opener))return{inTeams:!1,theme:"default",ssoToken:null};try{const{app:n,authentication:r}=await x(async()=>{const{app:l,authentication:h}=await import("./src-CZ2ouLw7.js");return{app:l,authentication:h}},[],import.meta.url);await Promise.race([n.initialize(),new Promise((l,h)=>setTimeout(()=>h(new Error("Teams init timeout")),350))]);const a=await n.getContext(),i=U(a.app?.theme);n.registerOnThemeChangeHandler(l=>{e(U(l))});let o=null;try{o=await r.getAuthToken()}catch{o=null}return{inTeams:!0,theme:i,ssoToken:o}}catch{return{inTeams:!1,theme:"default",ssoToken:null}}}const ae=[{id:"inbox",label:"Inbox"},{id:"sentitems",label:"Sent"},{id:"drafts",label:"Drafts"}],w=document.getElementById("root"),O=new Map;let t={theme:"default",inTeams:!1,auth:null,folder:"inbox",unreadOnly:!1,search:"",messages:[],selectedId:null,selectedDetail:null,loadingList:!1,loadingDetail:!1,replyOpen:!1,replyAll:!1,sendingReply:!1,composeOpen:!1,sendingCompose:!1,updatingApp:!1,versionInfo:null,error:null},A=null,R=null,L=null,k=null,I=null,p=[],f=-1;function d(e){return e?e.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&#39;"):""}function F(e){const s=new Date(e);return Number.isNaN(s.getTime())?e:s.toLocaleString([],{month:"short",day:"numeric",hour:"2-digit",minute:"2-digit"})}function j(e){t.theme=e,document.documentElement.setAttribute("data-theme",e)}function P(){return t.messages.length===0?`<div class="empty-state">${t.loadingList?"Loading messages…":"No messages found."}</div>`:t.messages.map(e=>{const s=e.from?.emailAddress.name||e.from?.emailAddress.address||"Unknown",n=e.isRead?"":"unread",r=t.selectedId===e.id?"selected":"",a=e.isRead?"":'<span class="unread-dot" aria-label="Unread"></span>';return`
        <button class="message-item ${n} ${r}" data-id="${d(e.id)}">
          <div class="message-item-top">
            <span class="sender-row">
              ${a}
              <span class="sender-name">${d(s)}</span>
            </span>
            <span class="message-time">${d(F(e.receivedDateTime))}</span>
          </div>
          <div class="message-subject">${d(e.subject||"(No subject)")}</div>
          <div class="message-preview">${d(e.bodyPreview)}</div>
        </button>`}).join("")}function V(){if(t.loadingDetail)return'<div class="empty-state">Loading message…</div>';const e=t.selectedDetail;if(!e)return'<div class="empty-state">Select a message to read or reply.</div>';const s=e.from?.emailAddress.name?`${e.from.emailAddress.name} <${e.from.emailAddress.address}>`:e.from?.emailAddress.address??"",n=e.body.contentType.toLowerCase()==="html"?e.body.content:`<pre class="email-body-text">${d(e.body.content)}</pre>`,r=t.replyOpen?`
      <div class="reply-box">
        <strong>${t.replyAll?"Reply All":"Reply"}</strong>
        <textarea id="reply-text" class="textarea" placeholder="Write your reply…"></textarea>
        <div class="detail-actions">
          <button id="send-reply-btn" class="btn btn-primary" ${t.sendingReply?"disabled":""}>
            ${t.sendingReply?"Sending…":"Send Reply"}
          </button>
          <button id="cancel-reply-btn" class="btn">Cancel</button>
        </div>
      </div>`:"";return`
    <div class="detail-card">
      <div class="detail-header">
        <div>
          <h2 class="detail-subject">${d(e.subject||"(No subject)")}</h2>
          <div class="detail-meta">
            <span><strong>From:</strong> ${d(s)}</span>
            <span><strong>Date:</strong> ${d(F(e.receivedDateTime))}</span>
          </div>
        </div>
        <div class="detail-actions">
          <button id="toggle-read-btn" class="btn">${e.isRead?"Mark unread":"Mark read"}</button>
          <button id="reply-btn" class="btn btn-primary">Reply</button>
          <button id="reply-all-btn" class="btn">Reply All</button>
        </div>
      </div>
      <div class="email-body">${n}</div>
      ${r}
    </div>`}function c(){const{auth:e}=t;if(!e){w.innerHTML='<div class="empty-state">Loading Outlook Mail…</div>';return}if(!e.authenticated){const m=e.entraConfigured?t.inTeams?'<button id="teams-login-btn" class="btn btn-primary">Sign in with Microsoft</button>':`<a class="btn btn-primary" href="${S}/api/auth/login">Sign in with Microsoft</a>`:"<p>Configure <code>AZURE_CLIENT_ID</code> and <code>AZURE_CLIENT_SECRET</code> in <code>.env</code>, or enable <code>DEV_MOCK_GRAPH=true</code>.</p>";w.innerHTML=`
      <div class="empty-state">
        <h2>Outlook Mail for Teams</h2>
        <p>Sign in with your Microsoft 365 account to view and manage your Outlook mailbox.</p>
        ${m}
      </div>`;return}const s=e.user?`<span class="user-label">${d(e.user.displayName)} (${d(e.user.mail)})</span>`:"",n=e.mode==="session"?'<button id="logout-btn" class="btn">Sign out</button>':"",r=t.error?`<div class="error-banner" role="alert"><span>${d(t.error)}</span><button id="dismiss-err" class="btn">Dismiss</button></div>`:"",a=ae.map(m=>`
      <button role="tab" aria-selected="${t.folder===m.id}"
        class="folder-tab ${t.folder===m.id?"active":""}"
        data-folder="${m.id}">${m.label}</button>`).join(""),i=t.composeOpen?`
      <div class="modal-backdrop" role="dialog" aria-modal="true" aria-label="New Mail">
        <form id="compose-form" class="modal">
          <h3>New Mail</h3>
          <div class="autocomplete-wrap">
            <input id="compose-to" type="text" class="field-input" placeholder="To (comma-separated emails)" autocomplete="off" required />
            <div id="compose-suggestions" class="autocomplete-list" hidden></div>
          </div>
          <input id="compose-subject" type="text" class="field-input" placeholder="Subject" required />
          <textarea id="compose-body" class="textarea" placeholder="Message body…"></textarea>
          <div class="modal-footer">
            <button type="button" id="cancel-compose-btn" class="btn">Cancel</button>
            <button type="submit" class="btn btn-primary" ${t.sendingCompose?"disabled":""}>
              ${t.sendingCompose?"Sending…":"Send"}
            </button>
          </div>
        </form>
      </div>`:"",o=t.versionInfo?`v${d(t.versionInfo.version)}`:"v1.0.4",l=t.versionInfo?`v${d(t.versionInfo.version)} (${d(t.versionInfo.commit)})`:"v1.0.4",h=t.versionInfo?.updateAvailable?`${t.versionInfo.behindCount} update(s) available`:"Up to date",u=t.versionInfo?.updateAvailable?'<span class="update-dot" title="Update available"></span>':"";w.innerHTML=`
    <header class="topbar">
      <div class="topbar-left">
        <h1 class="app-title">Outlook Mail</h1>
        <span class="mode-badge">${t.inTeams?"Teams Tab":"Browser"} · ${e.mode==="mock"?"Mock Graph":"Graph API"}</span>
        <div class="app-brand" tabindex="0">
          <span class="mode-badge">${o} ${u}</span>
          <div class="version-popover">
            <div class="version-card">
              <div><strong>Outlook Mail</strong> ${l}</div>
              <div class="version-meta">${d(h)}</div>
              <button id="app-update-btn" class="btn btn-primary" ${t.updatingApp?"disabled":""}>
                ${t.updatingApp?"Updating…":"Update"}
              </button>
            </div>
          </div>
        </div>
      </div>
      <div class="topbar-actions">
        ${s}
        <button id="new-mail-btn" class="btn btn-primary">New Mail</button>
        <button id="refresh-btn" class="btn" ${t.loadingList?"disabled":""}>
          ${t.loadingList?"Refreshing…":"Refresh"}
        </button>
        ${n}
      </div>
    </header>
    ${r}
    <div class="toolbar">
      <div class="folder-tabs" role="tablist" aria-label="Mail folders">${a}</div>
      <div class="search-controls">
        <label class="checkbox-label">
          <input id="unread-toggle" type="checkbox" ${t.unreadOnly?"checked":""} />
          Unread only
        </label>
        <input id="search-input" type="search" class="search-input" placeholder="Search mail…"
          aria-label="Search mail" value="${d(t.search)}" />
      </div>
    </div>
    <main class="workspace">
      <section id="message-list" class="message-list" aria-label="Messages">${P()}</section>
      <section id="detail-pane" class="detail-pane" aria-label="Message detail">${V()}</section>
    </main>
    ${i}`}function g(){const e=document.getElementById("message-list"),s=document.getElementById("detail-pane");if(!e||!s){c();return}e.innerHTML=P(),s.innerHTML=V();const n=document.getElementById("refresh-btn");n&&(n.disabled=t.loadingList,n.textContent=t.loadingList?"Refreshing…":"Refresh")}async function v(){A?.abort(),A=new AbortController,t.loadingList=!0,t.error=null,g();try{const e=await _({folder:t.folder,unreadOnly:t.unreadOnly,search:t.search},A.signal);t.messages=e.messages}catch(e){if(e.name==="AbortError")return;t.error=e instanceof Error?e.message:"Failed to load messages",c();return}finally{t.loadingList=!1}g()}async function oe(e){R?.abort(),R=new AbortController,t.selectedId=e,t.replyOpen=!1,t.error=null;const s=O.get(e);if(s)t.selectedDetail=s,t.loadingDetail=!1,g();else{t.loadingDetail=!0,g();try{const n=await Q(e,R.signal);O.set(e,n),t.selectedDetail=n}catch(n){if(n.name==="AbortError")return;t.error=n instanceof Error?n.message:"Failed to open message",t.loadingDetail=!1,c();return}finally{t.loadingDetail=!1}}t.selectedDetail&&!t.selectedDetail.isRead?(t.selectedDetail.isRead=!0,t.messages=t.messages.map(n=>n.id===e?{...n,isRead:!0}:n),g(),q(e,!0).catch(()=>{})):g()}w.addEventListener("click",async e=>{const s=e.target,n=s.closest(".message-item[data-id]");if(n?.dataset.id){oe(n.dataset.id);return}const r=s.closest(".folder-tab[data-folder]");if(r?.dataset.folder){t.folder=r.dataset.folder,t.selectedId=null,t.selectedDetail=null,c(),v();return}if(s.id==="refresh-btn")O.clear(),v();else if(s.id==="app-update-btn"){t.updatingApp=!0,c();try{await se();for(let a=0;a<20;a++)if(await new Promise(o=>setTimeout(o,500)),await fetch(`${S}/api/health`).then(o=>o.ok).catch(()=>!1)){window.location.reload();return}window.location.reload()}catch(a){t.updatingApp=!1,t.error=a instanceof Error?a.message:"Update failed",c()}}else if(s.closest(".autocomplete-item[data-email]")){const a=s.closest(".autocomplete-item[data-email]");z(a.dataset.email)}else if(s.id==="new-mail-btn")t.composeOpen=!0,c();else if(s.id==="cancel-compose-btn")t.composeOpen=!1,c();else if(s.id==="dismiss-err")t.error=null,c();else if(s.id==="teams-login-btn")try{const{authentication:a}=await x(async()=>{const{authentication:l}=await import("./src-CZ2ouLw7.js");return{authentication:l}},[],import.meta.url),i=await a.authenticate({url:`${S||window.location.origin}/api/auth/login?teams=1`,width:600,height:560});i&&N(i);const o=await C();t.auth=o,c(),o.authenticated&&v()}catch(a){t.error=a instanceof Error?a.message:"Microsoft sign-in failed",c()}else if(s.id==="logout-btn")await Z(),N(null),O.clear(),t.auth=await C(),t.messages=[],t.selectedDetail=null,c();else if(s.id==="toggle-read-btn"&&t.selectedDetail){const a=t.selectedDetail,i=!a.isRead;a.isRead=i,t.messages=t.messages.map(o=>o.id===a.id?{...o,isRead:i}:o),g();try{await q(a.id,i)}catch(o){t.error=o instanceof Error?o.message:"Failed to update read status",c()}}else if(s.id==="reply-btn")t.replyAll=!1,t.replyOpen=!0,g();else if(s.id==="reply-all-btn")t.replyAll=!0,t.replyOpen=!0,g();else if(s.id==="cancel-reply-btn")t.replyOpen=!1,g();else if(s.id==="send-reply-btn"&&t.selectedDetail){const i=document.getElementById("reply-text")?.value.trim()??"";if(!i)return;t.sendingReply=!0,g();try{await X(t.selectedDetail.id,i,t.replyAll),t.replyOpen=!1,t.sendingReply=!1,await v()}catch(o){t.sendingReply=!1,t.error=o instanceof Error?o.message:"Failed to send reply",c()}}});w.addEventListener("change",e=>{const s=e.target;s.id==="unread-toggle"&&(t.unreadOnly=s.checked,v())});function $(){const e=document.getElementById("compose-suggestions");if(e){if(p.length===0){e.hidden=!0,e.innerHTML="";return}e.hidden=!1,e.innerHTML=p.map((s,n)=>`
      <button type="button" class="autocomplete-item ${n===f?"active":""}" data-email="${d(s.address)}">
        <span><strong>${d(s.name)}</strong> &lt;${d(s.address)}&gt;</span>
        <span class="autocomplete-sub">${d(s.source)}</span>
      </button>`).join("")}}function z(e){const s=document.getElementById("compose-to");if(!s)return;const n=s.value.split(",");n[n.length-1]=` ${e}`,s.value=`${n.map(r=>r.trim()).filter(Boolean).join(", ")}, `,p=[],f=-1,$(),s.focus()}w.addEventListener("input",e=>{const s=e.target;if(s.id==="search-input")t.search=s.value,k&&window.clearTimeout(k),k=window.setTimeout(()=>void v(),200);else if(s.id==="compose-to"){const n=s.value.split(",").pop()?.trim()??"";if(I&&window.clearTimeout(I),n.length<1){p=[],f=-1,$();return}I=window.setTimeout(async()=>{L?.abort(),L=new AbortController;try{p=(await ee(n,L.signal)).suggestions,f=p.length>0?0:-1,$()}catch{}},150)}});w.addEventListener("keydown",e=>{e.target.id!=="compose-to"||p.length===0||(e.key==="ArrowDown"?(e.preventDefault(),f=(f+1)%p.length,$()):e.key==="ArrowUp"?(e.preventDefault(),f=(f-1+p.length)%p.length,$()):e.key==="Enter"&&f>=0?(e.preventDefault(),z(p[f].address)):e.key==="Escape"&&(p=[],f=-1,$()))});w.addEventListener("submit",async e=>{const s=e.target;if(s.id!=="compose-form")return;e.preventDefault();const n=document.getElementById("compose-to")?.value??"",r=document.getElementById("compose-subject")?.value.trim()??"",a=document.getElementById("compose-body")?.value??"",i=n.split(",").map(l=>l.trim()).filter(Boolean);if(i.length===0||!r)return;t.sendingCompose=!0;const o=s.querySelector('button[type="submit"]');o&&(o.disabled=!0,o.textContent="Sending…");try{await Y({to:i,subject:r,body:a}),t.composeOpen=!1,t.sendingCompose=!1,p=[],f=-1,c(),v()}catch(l){t.sendingCompose=!1,o&&(o.disabled=!1,o.textContent="Send"),t.error=l instanceof Error?l.message:"Failed to send email",c()}});async function H(){if(!(!t.auth?.authenticated||document.hidden||t.composeOpen||t.replyOpen||t.loadingList))try{const e=await _({folder:t.folder,unreadOnly:t.unreadOnly,search:t.search});t.messages=e.messages;const s=document.getElementById("message-list");s&&(s.innerHTML=P())}catch{}}const ie=3e4;async function re(){if(window.location.pathname==="/auth-done"){const i=new URLSearchParams(window.location.search).get("sid")??"";try{const{app:o,authentication:l}=await x(async()=>{const{app:h,authentication:u}=await import("./src-CZ2ouLw7.js");return{app:h,authentication:u}},[],import.meta.url);await o.initialize(),l.notifySuccess(i)}catch{window.close()}return}const e=await ne(j);t.inTeams=e.inTeams,j(e.theme),W(e.ssoToken);const s=C(),n=_({folder:t.folder,unreadOnly:t.unreadOnly,search:t.search}).catch(()=>null),[r,a]=await Promise.all([s,n]);t.auth=r,r.authenticated&&a?t.messages=a.messages:r.authenticated&&!a&&v(),c(),te(!0).then(i=>{t.versionInfo=i,c()}).catch(()=>{}),window.setInterval(()=>void H(),ie),document.addEventListener("visibilitychange",()=>{document.hidden||H()})}re();
