(function () {
    "use strict";

    var messageBanner;
    var ssoToken;
    var msgSubject;
    var mailMode;

    // The Office initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $(".loader").css("display", "block");
            $("#afailure").css("display", "none");
            getAccessToken();

            $("#drpcases").change((event) => {
                $("#savesection").css("display", "block");
            });

            $("#btnSave").click(function () {
                var mailRecepients = [{
                    "displayName": "",
                    "emailAddress": "ankerh@emails.itsm360cloud.net"
                }];

                var newSubject = msgSubject + "ID" + $("#drpcases").find("option:selected").val();
                Office.context.mailbox.item.subject.setAsync(newSubject, function (asyncResult) {
                    if (asyncResult.status === "failed") {
                        console.log("Action failed with error: " + asyncResult.error.message);
                        $("#afailure").text("Case ID not appended to the subject").css("display", "block");
                    } else {
                        console.log("Action Subject appended");
                        Office.context.mailbox.item.bcc.setAsync(mailRecepients, function (result) {
                            if (result.error) {
                                console.log(result.error);
                                $("#afailure").text("Failure while adding the bcc").css("display", "block");
                            } else {
                                console.log("Recipients added to the bcc");
                                Office.context.ui.closeContainer();
                            }
                        });
                    }
                });

            });
            
            var item = Office.context.mailbox.item;
            item.subject.getAsync(function (result) {
                msgSubject = result.value;
            });
        });
    };


    function getAccessToken() {
        if (Office.context.auth !== undefined && Office.context.auth.getAccessTokenAsync !== undefined) {
            Office.context.auth.getAccessTokenAsync(function (result) {
                if (result.status === "succeeded") {
                    console.log("token was fetched ");
                    ssoToken = result.value;
                    getCases(result.value);

                } else if (result.error.code === 13007 || result.error.code === 13005) {
                    console.log("fetching token by force consent");
                    Office.context.auth.getAccessTokenAsync({ forceConsent: true }, function (result) {
                        if (result.status === "succeeded") {
                            console.log("token was fetched");
                            ssoToken = result.value;
                            getCases(result.value);
                        }
                        else {
                            console.log("No token was fetched " + result.error.code);
                            //getSiteCollections();
                        }
                    });
                }
                else {
                    console.log("error while fetching access token " + result.error.code);
                    $(".loader").css("display", "none");
                }
            });
        }
    }

    function getCases(token) {

        $.ajax({
            type: "GET",
            url: "api/GetCases",
            headers: {
                "Authorization": "Bearer " + token
            },
            contentType: "application/json; charset=utf-8"
        }).done(function (data) {
            console.log("Fetched the Cases data");
            $.each(data, (index, value) => {
                $("#drpcases").append('<option value="' + value.ID + '">' + value.Title + '</option>');
            });
            $(".loader").css("display", "none");
        }).fail(function (error) {
            console.log("Fail to fetch cases");
            console.log(error);
            $(".loader").css("display", "none");
        });
    }
})();