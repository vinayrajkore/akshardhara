const fs = require('fs');
let htmlText = fs.readFileSync('index.html', 'utf8');

// 1. Fix mobile overlap and add margin to auth-page
// Replace padding-top of #auth-page.active and .auth-wrap margins
htmlText = htmlText.replace(
  /#auth-page\.active\s*\{[^}]*\}/g,
  "#auth-page.active{display:flex;flex-direction:column;min-height:100vh;padding:24px;background:var(--bg);padding-top:calc(var(--nav-h) + 40px) !important;}"
);

htmlText = htmlText.replace(
  /\.auth-wrap\s*\{\s*margin-top:\s*16px\s*!important;\s*margin-bottom:\s*32px\s*!important;\s*\}/g,
  ".auth-wrap { margin-top: 48px !important; margin-bottom: 32px !important; }"
);

// 2. Fix Superadmin frontend UI (because patch-superadmin.js failed to replace loadAUsers properly)
// We will replace the entire loadAUsers function to ensure it works.
const newLoadAUsers = `async function loadAUsers(){
  load(true);const r=await api({action:'getAllUsers',userID:CU.id});load(false);if(r.error){toast(r.error,'error');return;}
  document.getElementById('us-body').innerHTML=(r.users||[]).map(u=>{
    let roleBadge = '';
    if(u.role==='superadmin') roleBadge=\`<span class="badge badge-brown">\${I.shield} Super Admin</span>\`;
    else if(u.role==='admin' || u.role==='subadmin') roleBadge=\`<span class="badge badge-brown" style="opacity:0.85">\${I.shield} Sub Admin</span>\`;
    else roleBadge=\`<span class="badge badge-blue">\${I.user} Reader</span>\`;
    
    let actions = '';
    if(CU.role==='superadmin' && u.id!==CU.id) {
        if(u.role==='superadmin') actions='<span class="text-muted">Super Admin</span>';
        else {
            actions = \`<div class="flex gap-6">
              <button class="btn \${u.role==='admin' || u.role==='subadmin'?'btn-outline':'btn-gold'} btn-xs" onclick="toggleRole('\${u.id}','\${u.role==='admin' || u.role==='subadmin'?'reader':'subadmin'}')">\${u.role==='admin' || u.role==='subadmin'?t('make_reader'):t('make_admin')}</button>
              <button class="btn btn-danger btn-xs" onclick="deleteUser('\${u.id}')">\${t('delete')}</button>
            </div>\`;
        }
    } else if(u.id===CU.id) {
        actions='<span class="text-muted">You</span>';
    } else {
        actions='<span class="text-muted">-</span>';
    }

    return \`<tr>
      <td><strong>\${esc(u.name||'&mdash;')}</strong></td><td>\${esc(u.email)}</td><td>\${esc(u.mobile||'&mdash;')}</td><td>\${u.age||'&mdash;'}</td>
      <td>\${roleBadge}</td>
      <td>\${actions}</td>
    </tr>\`;
  }).join('');
}`;

const startIndex = htmlText.indexOf('async function loadAUsers(){');
const endIndex = htmlText.indexOf('async function deleteUser(uid){');

if (startIndex !== -1 && endIndex !== -1) {
  htmlText = htmlText.substring(0, startIndex) + newLoadAUsers + "\n\n" + htmlText.substring(endIndex);
} else {
  console.log("Could not find loadAUsers or deleteUser indices.");
}

fs.writeFileSync('index.html', htmlText, 'utf8');
console.log('Fixed index.html issues successfully!');
