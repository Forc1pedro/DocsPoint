import { NextResponse } from 'next/server';
import type { NextRequest } from 'next/server';

function applySecurityHeaders(response: NextResponse): NextResponse {
  response.headers.set('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  response.headers.set('X-Frame-Options', 'SAMEORIGIN');
  response.headers.set('X-Content-Type-Options', 'nosniff');
  response.headers.set('Referrer-Policy', 'no-referrer');
  return response;
}

// TASK-30: Next.js Middleware de Proteção de Rotas via Cookie HttpOnly sp_session e x-test-role
export function middleware(request: NextRequest) {
  const { pathname } = request.nextUrl;
  const spSession = request.cookies.get('sp_session')?.value;
  const testRole = request.headers.get('x-test-role') || request.headers.get('x-user-role');
  const isAuthenticated = Boolean(spSession || testRole);

  const isPublicRoute =
    pathname.startsWith('/login') ||
    pathname.startsWith('/shared') ||
    pathname.startsWith('/api') ||
    pathname.startsWith('/_next') ||
    pathname.startsWith('/assets') ||
    pathname === '/favicon.ico';

  // Se o usuário está na página de login mas já possui sessão ativa, redireciona para a home
  if (pathname === '/login' && spSession) {
    return applySecurityHeaders(NextResponse.redirect(new URL('/', request.url)));
  }

  // Se a rota for protegida e o usuário não estiver autenticado, redireciona para /login
  if (!isPublicRoute && !isAuthenticated) {
    const loginUrl = new URL('/login', request.url);
    return applySecurityHeaders(NextResponse.redirect(loginUrl));
  }

  return applySecurityHeaders(NextResponse.next());
}

export const config = {
  matcher: [
    '/((?!_next/static|_next/image|favicon.ico).*)',
  ],
};
