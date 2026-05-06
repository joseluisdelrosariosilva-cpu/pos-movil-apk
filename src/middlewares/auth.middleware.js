// Validación de token para rutas protegidas

import {
  haySesionActiva,
  sesionExpirada,
  renovarExpiracion,
  getSesion,
  esMismaIP,
} from "../utils/session.util.js";
import { validarFormatoToken } from "../utils/token.util.js";

// Rutas que NO requieren autenticación
// Nota: cuando el middleware se monta con app.use('/api', ...), req.path llega como '/productos',
// no '/api/productos'. Por eso incluimos ambas variantes.
const RUTAS_PUBLICAS = [
  "/",
  "/test",
  "/estado-publico",
  "/api/test",
  "/api/estado-publico",
];

// Rutas que permiten token vacío (para APK sin sesión)
const RUTAS_FLEXIBLES = [
  "/productos",
  "/ventas",
  "/mermas",
  "/resumen",
  "/api/productos",
  "/api/ventas",
  "/api/mermas",
  "/api/resumen",
  "/sync",
  "/api/sync",
  "/entrada-productos",
  "/api/entrada-productos",
];

const normalizarRuta = (ruta = "") => {
  if (!ruta) return "/";
  const limpia = ruta.split("?")[0];
  return limpia.endsWith("/") && limpia.length > 1
    ? limpia.slice(0, -1)
    : limpia;
};

/**
 * Middleware para validar token de sesión
 */
export const validarToken = (req, res, next) => {
  const rutaPath = normalizarRuta(req.path);
  const rutaOriginal = normalizarRuta(req.originalUrl);

  const coincideCon = (lista) =>
    lista.some((ruta) => ruta === rutaPath || ruta === rutaOriginal);

  // 1. Verificar si es ruta pública
  if (coincideCon(RUTAS_PUBLICAS)) {
    return next();
  }

  // 2. Obtener token del header
  const token = req.headers["x-session-token"];

  // 3. Rutas flexibles - permiten token vacío o sin sesión (para APK)
  if (coincideCon(RUTAS_FLEXIBLES)) {
    // Si no hay sesión activa ni token, permitir acceso (fallback offline)
    if (!haySesionActiva() && !token) {
      console.log("🔓 [auth] Ruta flexible sin sesión - permitiendo:", rutaPath);
      return next();
    }
    // Si hay sesión pero el token no es válido exactamente, verificar si es IP válida
    if (haySesionActiva()) {
      // Permitir si hay sesión activa (no verificamos token exactamente para rutas flexibles)
      console.log("🔓 [auth] Ruta flexible con sesión activa - permitiendo:", rutaPath);
      return next();
    }
    // Si hay token pero sesión inactive, verificar formato básico
    if (token && validarFormatoToken(token)) {
      console.log("🔓 [auth] Ruta flexible con token válido - permitiendo:", rutaPath);
      return next();
    }
  }

  // 4. Verificar si hay sesión activa (para rutas protegidas)
  if (!haySesionActiva()) {
    console.log("⛔ Petición sin sesión activa:", rutaPath);
    return res.status(401).json({
      error: "NO_HAY_SESION",
      mensaje: "No hay una sesión activa. Visita la página principal.",
    });
  }

  // 5. Verificar expiración
  if (sesionExpirada()) {
    console.log("⛔ Sesión expirada");
    return res.status(401).json({
      error: "SESION_EXPIRADA",
      mensaje: "La sesión ha expirado. Vuelve a la página principal.",
    });
  }

  // 5. Validar formato del token
  if (!validarFormatoToken(token)) {
    console.log("⛔ Formato de token inválido");
    return res.status(403).json({
      error: "TOKEN_INVALIDO",
      mensaje: "Formato de token inválido.",
    });
  }

  // 6. Validar que el token coincide
  const { tokenActivo, ipActiva } = getSesion();

  if (token !== tokenActivo) {
    console.log(
      "⛔ Token no coincide. Esperado:",
      tokenActivo,
      "Recibido:",
      token,
    );
    return res.status(403).json({
      error: "TOKEN_INCORRECTO",
      mensaje: "Token de sesión incorrecto.",
    });
  }

  // 7. Validar que la IP coincide (seguridad extra)
  const ipCliente = req.ip;
  if (!esMismaIP(ipCliente)) {
    console.log("⛔ IP diferente:", ipCliente, "vs", ipActiva);
    return res.status(403).json({
      error: "IP_DIFERENTE",
      mensaje: "La IP no coincide con la sesión activa.",
    });
  }

  // 8. Todo OK - renovar expiración
  renovarExpiracion();

  // 9. Continuar
  next();
};
