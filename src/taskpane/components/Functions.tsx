

  // export var salutation = function(emailTo: Office.Recipients) {
  export var salutation = function(emailTo) {
  // export function salutation() {

    // let emailTo: Office.Recipients = Office.context.mailbox.item.to;

     var salutation: string = "";
     var fullName: string = "";
     var firstName: string = "";
     var countOfTo:number = 0;
     var toEnumerator:number = 0;
     countOfTo  = emailTo.lastIndex;

    //  for (let toEnumerator = 0, toEnumerator < countOfTo; toEnumerator++) {
     for (toEnumerator = 0; toEnumerator <= countOfTo; toEnumerator) {
      fullName =  emailTo[toEnumerator++]["displayName"];
      firstName = fullName.split(" ", 1).toString(); 
       salutation = salutation + " " + firstName + " and";
     }
     
     //remove last "and"
     var salutationLength:number = 0;
     salutationLength = salutation.length;

     salutation = salutation.substr(1,salutationLength - 5);
     
     salutation = "Hi " + salutation;
     
     return salutation;

  };
