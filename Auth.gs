function _hash_(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function _isAdmin_(user) {
  const r = (user.r || "").toString().toUpperCase();
  return r === "ADMIN" || r === "ADMINISTRADOR";
}

function login(usuario, password) {
  usuario = (usuario || "").toString().trim();
  password = (password || "").toString();

  if (!usuario || !password) return { ok:false, msg:"Completá usuario y contraseña." };

  const sh = _sheet(SHEET_USERS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return { ok:false, msg:"No hay usuarios cargados." };

  const headers = data[0].map(x => (x || "").toString().trim());
  const idx = (name) => headers.indexOf(name);

  const iUsuario = idx("Usuario");
  const iPass = idx("Password");
  const iNombre = idx("Nombre");
  const iRol = idx("Rol");
  const iActivo = idx("Activo");

  if ([iUsuario,iPass,iNombre,iRol,iActivo].some(i => i === -1)) {
    return { ok:false, msg:"La pestaña Usuarios no tiene los encabezados esperados." };
  }

  const row = data.slice(1).find(r => (r[iUsuario] || "").toString().trim().toLowerCase() === usuario.toLowerCase());
  if (!row) return { ok:false, msg:"Usuario o contraseña incorrectos." };

  const activo = (row[iActivo] || "").toString().trim().toLowerCase();
  if (!(activo === "si" || activo === "sí" || activo === "true" || activo === "1" || activo === "activo")) {
    return { ok:false, msg:"Usuario inactivo." };
  }

  const stored = (row[iPass] || "").toString();
  const passOk = (() => {
    if (stored.startsWith("sha256:")) {
      const hash = stored.replace("sha256:", "").trim().toLowerCase();
      return _hash_(password).toLowerCase() === hash;
    }
    return stored === password;
  })();

  if (!passOk) return { ok:false, msg:"Usuario o contraseña incorrectos." };

  const token = Utilities.getUuid();
  const payload = {
    u: usuario,
    n: (row[iNombre] || "").toString(),
    r: (row[iRol] || "").toString().trim().toUpperCase(),
    ts: Date.now()
  };

  CacheService.getScriptCache().put("SESS_" + token, JSON.stringify(payload), 60 * 60 * 6);
  return { ok:true, token, user: payload };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove("SESS_" + token);
  return { ok:true };
}

function getSession(token) {
  if (!token) return { ok:false };
  const raw = CacheService.getScriptCache().get("SESS_" + token);
  if (!raw) return { ok:false };
  return { ok:true, user: JSON.parse(raw) };
}

function _requireSession_(token) {
  const s = getSession(token);
  if (!s.ok) throw new Error("Sesión vencida. Volvé a iniciar sesión.");
  return s.user;
}
