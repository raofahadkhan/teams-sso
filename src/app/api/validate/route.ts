import { NextResponse } from 'next/server';
import jwt from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';

// Initialize the JWKS client
const client = jwksClient({
  jwksUri: 'https://login.microsoftonline.com/common/discovery/v2.0/keys',
});

// Helper function to retrieve the signing key
async function getSigningKey(header: any) {
  return new Promise((resolve, reject) => {
    client.getSigningKey(header.kid, (err, key) => {
      if (err) {
        return reject(err);
      }
      resolve(key!.getPublicKey());
    });
  });
}

// API route handler for validating tokens
export async function POST(req: Request) {
  try {
    // Extract the authorization token from the request headers
    const authorizationHeader = req.headers.get('authorization');
    const token = authorizationHeader?.split(' ')[1];

    if (!token) {
      return NextResponse.json({ error: 'No token provided' }, { status: 401 });
    }

    // Decode and verify the token
    const decoded = await new Promise((resolve, reject) => {
      jwt.verify(
        token,
        async (header: any, callback: (err: any, signingKey: string) => void) => {
          try {
            const signingKey = await getSigningKey(header) as string;
            callback(null, signingKey);
          } catch (error) {
            callback(error, '');
          }
        },
        { algorithms: ['RS256'] },
        (err, decodedToken) => {
          if (err) {
            return reject(err);
          }
          resolve(decodedToken);
        }
      );
    });

    return NextResponse.json({ decoded }, { status: 200 });
  } catch (error) {
    return NextResponse.json({ error: 'Invalid token', message: error }, { status: 401 });
  }
}
