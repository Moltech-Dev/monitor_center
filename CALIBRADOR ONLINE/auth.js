/* Acceso compartido con Monitor Center. Requiere el mismo origen y la misma pestaña. */
window.MCCalibradorAuth = {
  hasAccess() {
    const roles = (sessionStorage.getItem('ROLE') || '').toLowerCase().split(/[;,]/).map(role => role.trim());
    return Boolean(sessionStorage.getItem('USER')) && (roles.includes('admin') || roles.includes('calibrador'));
  },
  guard() {
    if (!sessionStorage.getItem('USER')) {
      window.location.replace('../index.html');
    } else if (!this.hasAccess()) {
      window.location.replace('../home.html');
    }
  },
  logout() {
    sessionStorage.clear();
    window.location.href = '../index.html';
  }
};
