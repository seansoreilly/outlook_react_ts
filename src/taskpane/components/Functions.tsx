export var salutation = function (emailTo) {

  var salutation: string = "";
  var countOfTo: number = 0;
  countOfTo = emailTo.lastIndex;

  switch (countOfTo) {
    case 0: {
      console.log(countOfTo);
      salutation =
        "Hi " + firstName(emailTo[0]["displayName"]);
      break;
    }
    case 1: {
      console.log(countOfTo);
      salutation =
        "Hi "
        + firstName(emailTo[0]["displayName"]) + " and "
        + firstName(emailTo[1]["displayName"]);
      break;
    }
    case 2: {
      console.log(countOfTo);
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
