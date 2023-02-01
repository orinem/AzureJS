const myMSALObj = new msal.PublicClientApplication(msalConfig);

let accountId = "";

function handleResponse(resp) {
  if (resp !== null) {
      console.log('id_token acquired at: ' + new Date().toString());
      console.log(resp)
      accountId = resp.account.homeAccountId;
      myMSALObj.setActiveAccount(resp.account);
      showWelcomeMessage(resp.account);
  } else {
      // need to call getAccount here?
      const currentAccounts = myMSALObj.getAllAccounts();
      if (!currentAccounts || currentAccounts.length < 1) {
          return;
      } else if (currentAccounts.length > 1) {
          // Add choose account code here
      } else if (currentAccounts.length === 1) {
          const activeAccount = currentAccounts[0];
          myMSALObj.setActiveAccount(activeAccount);
          accountId = activeAccount.homeAccountId;
          showWelcomeMessage(activeAccount);
      }
  }
}

async function signIn() {
  return myMSALObj.loginPopup(loginRequest).then(handleResponse).catch(error => {
      console.log(error);
    });
}

function signOut() {
  const logoutRequest = {
    account: myMSALObj.getAccountByHomeId(accountId)
  };
  myMSALObj.logoutPopup(logoutRequest).then(() => {
    window.location.reload();
  });
}

function callMSGraph(theUrl, accessToken, callback) {
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.onreadystatechange = function () {
        if (this.readyState == 4 && this.status == 200) {
           callback(JSON.parse(this.responseText));
        }
    }
    xmlHttp.open("GET", theUrl, true); // true for asynchronous
    xmlHttp.setRequestHeader('Authorization', 'Bearer ' + accessToken);
    xmlHttp.send();
}

async function getTokenPopup(request, account) {
  return await myMSALObj.acquireTokenSilent(request).catch(async (error) => {
      console.log("silent token acquisition fails.");
      if (error instanceof msal.InteractionRequiredAuthError) {
          console.log("acquiring token using popup");
          return myMSALObj.acquireTokenPopup(request).catch(error => {
              console.error(error);
          });
      } else {
          console.error(error);
      }
  });
}

async function seeProfile() {
  const currentAcc = myMSALObj.getAccountByHomeId(accountId);
  if (currentAcc) {
    const response = await getTokenPopup(loginRequest, currentAcc).catch(error => {
      console.log(error);
    });
    callMSGraph(graphConfig.graphMeEndpoint, response.accessToken, updateUI);
    profileButton.classList.add('d-none');
    mailButton.classList.remove('d-none');
  }
}

async function readMail() {
  const currentAcc = myMSALObj.getAccountByHomeId(accountId);
  if (currentAcc) {
      const response = await getTokenPopup(tokenRequest, currentAcc).catch(error => {
          console.log(error);
      });
      callMSGraph(graphConfig.graphMailEndpoint, response.accessToken, updateUI);
      mailButton.style.display = 'none';
  }
}
