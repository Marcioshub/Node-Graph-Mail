// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE.txt in the project root for license information.
var express = require("express");
var router = express.Router();
var authHelper = require("../helpers/auth");
var graph = require("@microsoft/microsoft-graph-client");

/* GET /mail */
router.get("/", async function(req, res, next) {
  let parms = { title: "Inbox", active: { inbox: true } };

  const accessToken = await authHelper.getAccessToken(req.cookies, res);
  const userName = req.cookies.graph_user_name;

  if (accessToken && userName) {
    parms.user = userName;

    // Initialize Graph client
    const client = graph.Client.init({
      authProvider: done => {
        done(null, accessToken);
      }
    });

    try {
      // Get the 10 newest messages from inbox
      const result = await client
        .api("/me/mailfolders/inbox/messages")
        .top(100)
        .header("Prefer", 'outlook.body-content-type="text"')
        .select("subject,from,receivedDateTime,isRead,bodyPreview,uniqueBody")
        .orderby("receivedDateTime DESC")
        .get();

      for (var i = 0; i < 100; i++) {
        result.value[i].uniqueBody.content = stripHtml(
          result.value[i].uniqueBody.content
        );

        if (
          result.value[i].uniqueBody.content.includes("marcio") ||
          result.value[i].uniqueBody.content.includes("Marcio")
        ) {
          result.value[i].uniqueBody.hasName = false;
        } else {
          result.value[i].uniqueBody.hasName = true;
        }
      }

      parms.messages = result.value;
      //console.log(result.value);
      res.render("mail", parms);
      //res.json(result.value);
    } catch (err) {
      parms.message = "Error retrieving messages";
      parms.error = { status: `${err.code}: ${err.message}` };
      parms.debug = JSON.stringify(err.body, null, 2);
      //res.render("error", parms);
      res.json({
        error: parms.message
      });
    }
  } else {
    // Redirect to home
    res.redirect("/");
  }
});

// remove some of the html crap
function stripHtml(html) {
  return checkForMyName(html.replace(/<[^>]*>?/gm, ""));
}

// check if my name is inside the email body
function checkForMyName(html) {
  var newtext = html.toLowerCase();
  //console.log("NewText:", newtext.value.uniqueBody.content.hasName);

  return newtext;
}

module.exports = router;
