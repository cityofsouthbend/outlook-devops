function idURL($) {
  return `https://dev.azure.com/southbendin/_apis/wit/workitems?ids=${$.val}&api-version=6.0`;
}

function ticketURL($) {
  return `https://dev.azure.com/southbendin/Applications%20-%20Project%20Portfolio/_apis/wit/workitems/$${$.val}?bypassrules=true&api-version=6.0`;
}

export { idURL, ticketURL };
