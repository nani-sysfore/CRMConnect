/// <reference path="adal.js" />

var user, authContext, errorMessage;
var organizationURI = "https://cirrusdemo.crm11.dynamics.com";
(function () {
    var tenant = "cirrusdemo.onmicrosoft.com"; 
    var clientId = "13affba6-c028-4974-a8a7-41faf1a1ef9e"; 
    var pageUrl = "http://crmconnector.azurewebsites.net/";

    var endpoints = {
        orgUri: organizationURI
    };

    window.config = {
        tenant: tenant,
        clientId: clientId,
        postLogoutRedirectUri: pageUrl,
        endpoints: endpoints,
        cacheLocation: 'localStorage', 
    };
    authContext = new AuthenticationContext(config);
    authenticate();
    document.getElementById('login').addEventListener('click', function () {
        login();
    })
    document.getElementById('btn_get_name').addEventListener('click', function () {
        authContext.acquireToken(organizationURI, getUserId)
    })
    document.getElementById('sign_out').addEventListener('click', function () {
        authContext.logOut();
    })

})();

function authenticate() {
    var isCallback = authContext.isCallback(window.location.hash);
    if (isCallback) {
        authContext.handleWindowCallback();
    }
    var loginError = authContext.getLoginError();

    if (isCallback && !loginError) {
        window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
    }
    else {
        //errorMessage.textContent = loginError;
        //alert(loginError);
    }
    user = authContext.getCachedUser();
    if(user)
    {
        displayLogin();
    }
}
function login() {
    authContext.login();
    
}

function displayLogin() {
    var anonymous_div = document.getElementById('anonymous_user')
    anonymous_div.style.display = 'none';

    document.getElementById('register_user').style.display = 'block';

    var helloMessage = document.createElement("span");
    helloMessage.textContent = "Hello " + user.profile.name;
    document.getElementById('user_name').appendChild(helloMessage);   
    
}

function getUserId(error,token) {
    var req = new XMLHttpRequest
    req.open("GET", encodeURI(organizationURI + "/api/data/v8.0/WhoAmI"), true);
    req.onreadystatechange = function () {
        if (req.readyState == 4 && req.status == 200) {
            var whoAmIResponse = JSON.parse(req.responseText);
            getFullname(whoAmIResponse.UserId)
        }
    };
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Authorization", "Bearer " + token);
    req.send();
}

function getFullname(Id) {
    var req = new XMLHttpRequest
    req.open("GET", encodeURI(organizationURI + "/api/data/v8.0/systemusers(" + Id + ")?$select=fullname"), true);
    req.onreadystatechange = function () {
        if (req.readyState == 4 && req.status == 200) {
            var userInfoResponse = JSON.parse(req.responseText);
            alert(userInfoResponse.fullname);
        }
    };
    req.setRequestHeader("OData-MaxVersion", "4.0");
    req.setRequestHeader("OData-Version", "4.0");
    req.setRequestHeader("Accept", "application/json");
    req.setRequestHeader("Authorization", "Bearer " + authContext.getCachedToken(organizationURI));
    req.send();
}


