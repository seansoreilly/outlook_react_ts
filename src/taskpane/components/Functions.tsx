export var salutation = function (emailTo:any) {
  var salutation: string = "";
  var countOfTo: number = 0;
  countOfTo = emailTo.lastIndex;
// type: Property 'lastIndex' does not exist on type 'EmailAddressDetails'

  switch (countOfTo) {
    case 0: {
      salutation =
        "Hi " + firstName(emailTo[0]["displayName"]);
      break;
    }
    case 1: {
      salutation =
        "Hi "
        + firstName(emailTo[0]["displayName"]) + " and "
        + firstName(emailTo[1]["displayName"]);
      break;
    }
    case 2: {
      salutation =
        "Hi "
        + firstName(emailTo[0]["displayName"]) + ", "
        + firstName(emailTo[1]["displayName"]) + " and "
        + firstName(emailTo[2]["displayName"])
      break;
    }
    default: {
      salutation =
        "Hi all";
      break;
    }
  }

  return salutation;

};

function firstName(fullName: string) {
  return fullName.split(" ", 1).toString();

}
