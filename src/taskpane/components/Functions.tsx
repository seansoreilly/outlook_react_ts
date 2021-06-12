

  export var salutation = function(emailTo: Office.Recipients) {
  // export function salutation() {

    // let emailTo: Office.Recipients = Office.context.mailbox.item.to;

     let salutation: string = "";
     let firstName: string = "";
     let counter: number = 0;

     for (var t in emailTo) {
    //  for (let t in emailTo["displayName"]) {
      let fullName =  emailTo[counter]["displayName"];
      firstName = fullName.split(" ", 1).toString(); 
       salutation = salutation + firstName + " and";
     }
     
     //remove last "and"
     salutation = salutation.substr(0,salutation.length - 3);
     
     salutation = "Hi " + salutation;
     
     console.log(salutation);

     return salutation;

  };
