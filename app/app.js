const msalConfig = {
  auth: {
    clientId: '1d7c3bbf-7491-4a65-b78c-673d0938a357',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: 'http://localhost:3000',
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  },
}

const loginRequest = { scopes: ['openid', 'profile', 'User.Read'] }

const tokenRequest = { scopes: ['User.Read', 'Mail.Read'] }

const graphConfig = {
  graphMeEndpoint: 'https://graph.microsoft.com/v1.0/me',
  graphMailEndpoint: 'https://graph.microsoft.com/v1.0/me/messages',
}

const msalInstance = new msal.PublicClientApplication(msalConfig)

let username = ''

function loadPage() {
  const currentAccounts = msalInstance.getAllAccounts()

  if (currentAccounts === null) {
    return
  } else if (currentAccounts.length > 1) {
    console.warn('Multiple accounts detected.')
  } else if (currentAccounts.length === 1) {
    username = currentAccounts[0].username
    welcome(currentAccounts[0])
  }
}

function handleResponse(resp) {
  if (resp !== null) {
    username = resp.account.username
    welcome(resp.account)
  } else {
    loadPage()
  }
}

function signIn() {
  msalInstance
    .loginPopup(loginRequest)
    .then(handleResponse)
    .catch(console.error)
}

function signOut() {
  const logoutRequest = {
    accounts: msalInstance.getAccountByUsername(username),
  }
  msalInstance.logout(logoutRequest)
}

function getTokenPopup(request) {
  request.account = msalInstance.getAccountByUsername(username)
  return msalInstance.acquireTokenSilent(request).catch((error) => {
    console.warn(
      'silent token acquisition fails. acquiring token using redirect'
    )
    if (error instanceof msal.InteractionRequiredAuthError) {
      return msalInstance
        .acquireTokenPopup(request)
        .then((tokenResponse) => {
          console.log(tokenResponse)

          return tokenResponse
        })
        .catch(console.error)
    } else {
      console.warn(error)
    }
  })
}

function showProfile() {
  getTokenPopup(loginRequest)
    .then((response) => {
      callMSGraph(
        graphConfig.graphMeEndpoint,
        response.accessToken,
        showProfileUI
      )
    })
    .catch(console.error)
}

function showMail() {
  getTokenPopup(tokenRequest)
    .then((response) => {
      console.log('got token', response.accessToken)
      callMSGraph(
        graphConfig.graphMailEndpoint,
        response.accessToken,
        showMailUI
      )
    })
    .catch(console.error)
}

function callMSGraph(endpoint, token, callback) {
  const headers = new Headers()
  const bearer = `Bearer ${token}`

  headers.append('Authorization', bearer)

  const options = { method: 'GET', headers }

  console.log(`request made to Graph API at: ${new Date().toString()}`)

  fetch(endpoint, options)
    .then((response) => response.json())
    .then((json) => callback(json, endpoint))
    .catch(console.error)
}

function welcome(account) {
  const signInButton = document.querySelector('#sign-in')
  const welcome = document.querySelector('#welcome')

  signInButton.classList.add('hidden')
  welcome.classList.remove('hidden')
  welcome.querySelector('#username').innerHTML = account.username
}

function showProfileUI(data) {
  console.log('profile data', data)
  const profileButton = document.querySelector('#profile-button')
  const profileContent = document.querySelector('#profile-content')

  profileContent.querySelector('.title').innerHTML = `${data.jobTitle}`
  profileContent.querySelector('.mail').innerHTML = `${data.mail}`
  profileContent.querySelector('.phone').innerHTML = `${data.businessPhones[0]}`
  profileContent.querySelector('.location').innerHTML = `${data.officeLocation}`
  profileButton.classList.add('hidden')
  profileContent.classList.remove('hidden')
}

function showMailUI(data) {
  console.log('profile data', data)
  const mailButton = document.querySelector('#mail-button')
  const mailContent = document.querySelector('#mail-content')

  if (data.value.length < 1) {
    mailContent.querySelector('.empty').classList.remove('hidden')
    mailContent.querySelector('.list').classList.add('hidden')
  } else {
    data.value.slice(0, 10).map((d, i) => {
      const item = document.createElement('li')

      item.innerHTML = d.subject
      mailContent.querySelector('.list').appendChild(item)
    })
  }

  mailButton.classList.add('hidden')
  mailContent.classList.remove('hidden')
}

loadPage()
