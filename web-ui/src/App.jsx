import { useState } from 'react'
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from '@azure/msal-react'
import { InteractionRequiredAuthError } from '@azure/msal-browser'
import { loginRequest, apiConfig } from './authConfig'

function SignInButton() {
  const { instance } = useMsal()
  return (
    <button onClick={() => instance.loginRedirect(loginRequest)}>
      Sign in with Microsoft
    </button>
  )
}

function SignOutButton() {
  const { instance } = useMsal()
  return (
    <button onClick={() => instance.logoutRedirect()}>
      Sign out
    </button>
  )
}

function Profile() {
  const { accounts } = useMsal()
  const account = accounts[0]
  return (
    <div>
      <p><strong>Name:</strong> {account?.name}</p>
      <p><strong>Username:</strong> {account?.username}</p>
    </div>
  )
}

function ApiProfile() {
  const { instance, accounts } = useMsal()
  const [profile, setProfile] = useState(null)
  const [error, setError] = useState(null)
  const [loading, setLoading] = useState(false)

  async function callApi() {
    setLoading(true)
    setError(null)
    try {
      let tokenResult
      try {
        tokenResult = await instance.acquireTokenSilent({
          ...loginRequest,
          account: accounts[0],
        })
      } catch (err) {
        if (err instanceof InteractionRequiredAuthError) {
          tokenResult = await instance.acquireTokenPopup(loginRequest)
        } else {
          throw err
        }
      }

      const response = await fetch(apiConfig.profileEndpoint, {
        headers: { Authorization: `Bearer ${tokenResult.accessToken}` },
      })
      if (!response.ok) {
        throw new Error(`API request failed: ${response.status} ${response.statusText}`)
      }
      setProfile(await response.json())
    } catch (err) {
      setError(err.message)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div style={{ marginTop: 24 }}>
      <button onClick={callApi} disabled={loading}>
        {loading ? 'Calling API…' : 'Call API'}
      </button>
      {error && <p style={{ color: 'crimson' }}>{error}</p>}
      {profile && (
        <pre style={{ background: '#f4f4f4', padding: 12, overflowX: 'auto' }}>
          {JSON.stringify(profile, null, 2)}
        </pre>
      )}
    </div>
  )
}

export default function App() {
  return (
    <div style={{ fontFamily: 'sans-serif', maxWidth: 480, margin: '80px auto', padding: '0 16px' }}>
      <h1>Entra ID OIDC Sample</h1>
      <AuthenticatedTemplate>
        <p>Signed in successfully.</p>
        <Profile />
        <ApiProfile />
        <SignOutButton />
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <p>Sign in to continue.</p>
        <SignInButton />
      </UnauthenticatedTemplate>
    </div>
  )
}
