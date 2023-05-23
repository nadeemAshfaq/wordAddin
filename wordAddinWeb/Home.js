var app = angular.module('myApp', ['ngMaterial', 'ngRoute']);

app.controller('myAppCtrl', function ($scope) {

    Office.initialize = function () {


        $scope.valueset = function () {

            Word.run(function (context) {
                const body = context.document.body;
                context.load(body, 'text');
               

                return context.sync().then(function () {
                    var bodyText = body.text;
                    console.log("Body Text: " + bodyText);

                    // Store the value in local storage
                    localStorage.setItem('bodyText', bodyText);
                    console.log("Value stored in local storage.");

                    
                });
            });




        };






   



        $scope.myFunction = function () {

            Word.run(function (context) {

                const body = context.document.body;
                context.load(body, 'text');

                return context.sync().then(function () {
                    var bodyText = body.text;
                    console.log("Body Text: " + bodyText);

                    // Retrieve the stored value from local storage
                    var storedText = localStorage.getItem('bodyText');
                    console.log("Value retrieved from local storage: " + storedText);

                    // Find the matching text within the body
                    var matchIndex = bodyText.indexOf(storedText);
                    if (matchIndex !== -1) {
                        console.log("Match found at index: " + matchIndex);
                        var matchText = bodyText.substr(matchIndex, storedText.length);
                        console.log("Matching text: " + matchText);
                        $scope.matchText = matchText;
                    } else {
                        console.log("No match found.");
                        $scope.matchText = ""; // Clear the value if no match is found
                    }

                    // Update the AngularJS scope
                    $scope.$apply();
                });

            
            });

        
       
       
        };

        $scope.removeValueFromLocalStorage = function () {
            localStorage.removeItem('bodyText');
            console.log("Value removed from local storage.");
        };




    };
   
})