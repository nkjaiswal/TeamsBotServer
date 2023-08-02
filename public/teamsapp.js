(function () {
    "use strict";
  
    // Call the initialize API first
    microsoftTeams.app.initialize().then(function () {
      microsoftTeams.app.getContext().then(function (context) {
        if (context?.app?.host?.name) {
          updateHubState(context.app.host.name);
        }
        if(context?.user) {
          updateWelcomeMessage(context.user);
        }
      });
    });
  
    function updateHubState(hubName) {
      if (hubName) {
        document.getElementById("hubState").innerHTML = "App: " + hubName;
      }
    }
    function updateWelcomeMessage(user) {
      if (user) {
        document.getElementById("welcomemessage").innerHTML = "Welcome " + user.userPrincipalName;
      }
    }
  })();
  