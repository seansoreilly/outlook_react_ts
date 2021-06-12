

  export var salutation = function() {

    let emailTo: Office.Recipients = Office.context.mailbox.item.to;

     let salutation: string = "";
     let firstName: string = "";

     for (let t in emailTo["displayName"]) {
      console.log(t);
      firstName = t.split(" ", 1).toString(); 
       salutation = salutation + firstName + " and";
     }
     
     //remove last "and"
     salutation = salutation.substr(0,salutation.length - 3);
     
     salutation = "Hi " + salutation;
     
     console.log(salutation);


  };
