import { NextRequest, NextResponse } from 'next/server';

// TASK-30: Next.js API Proxy interno para backend Express
async function handleProxy(req: NextRequest, { params }: { params: { path: string[] } }) {
  const subPath = params.path ? params.path.join('/') : '';
  const search = req.nextUrl.search || '';

  const spSession = req.cookies.get('sp_session')?.value;

  const forwardHeaders = new Headers();
  req.headers.forEach((value, key) => {
    const lowerKey = key.toLowerCase();
    if (!['host', 'connection', 'content-length'].includes(lowerKey)) {
      forwardHeaders.set(key, value);
    }
  });

  if (spSession) {
    forwardHeaders.set('Authorization', `Bearer ${spSession}`);
  }

  const isBodyAllowed = !['GET', 'HEAD'].includes(req.method);
  const body = isBodyAllowed ? await req.arrayBuffer() : undefined;

  const internalUrl = `${process.env.INTERNAL_API_URL || 'http://backend:5000/api/v1'}/${subPath}${search}`;
  const fallbackUrl = `http://localhost:5000/api/v1/${subPath}${search}`;

  let backendResponse: Response;
  try {
    backendResponse = await fetch(internalUrl, {
      method: req.method,
      headers: forwardHeaders,
      body,
      redirect: 'manual',
    });
  } catch {
    backendResponse = await fetch(fallbackUrl, {
      method: req.method,
      headers: forwardHeaders,
      body,
      redirect: 'manual',
    });
  }

  const responseHeaders = new Headers();
  backendResponse.headers.forEach((value, key) => {
    if (key.toLowerCase() === 'set-cookie') {
      responseHeaders.append(key, value);
    } else {
      responseHeaders.set(key, value);
    }
  });

  const responseBody = await backendResponse.arrayBuffer();

  return new NextResponse(responseBody, {
    status: backendResponse.status,
    statusText: backendResponse.statusText,
    headers: responseHeaders,
  });
}

export const GET = handleProxy;
export const POST = handleProxy;
export const PUT = handleProxy;
export const DELETE = handleProxy;
export const PATCH = handleProxy;
export const HEAD = handleProxy;
