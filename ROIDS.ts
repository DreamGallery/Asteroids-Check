interface Query {
  query: string;
  variables: Record<string, unknown>;
}

interface ResponseAtom {
  balances: Array<{
    denom: string;
    amount: string;
  }>;
  pagination: {
    next_key: null;
    total: number;
  };
}

interface ResponseRoids {
  data: {
    token_holder: Array<{
      token: {
        ticker: string;
        content_path: string;
        max_supply: number;
        circulating_supply: number;
        decimals: number;
        last_price_base: number;
        transaction: {
          hash: string;
        };
      };
      amount: number;
      date_updated: string;
    }>;
  };
}

interface ResponseList {
  data: {
    token_open_position: Array<{
      id: number;
      ppt: number;
      amount: number;
      total: number;
    }>;
  };
}

async function postJson(
  url: string,
  data: Query
): Promise<ResponseRoids | ResponseList | string> {
  try {
    const options: RequestInit = {
      method: "post",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(data),
    };
    const response = await fetch(url, options);
    const content = await response.text();
    const json = JSON.parse(content) as ResponseRoids | ResponseList;
    return json;
  } catch (err) {
    if (err instanceof Error) {
      return "Error posting data: " + err.message;
    } else {
      return "Unknown error: " + err;
    }
  }
}

async function getJson(url: string): Promise<ResponseAtom | string> {
  try {
    const options: RequestInit = {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
      },
    };
    const response = await fetch(url, options);
    const content = await response.text();
    const json = JSON.parse(content) as ResponseAtom;
    return json;
  } catch (err) {
    if (err instanceof Error) {
      return "Error posting data: " + err.message;
    } else {
      return "Unknown error: " + err;
    }
  }
}

async function getRoidsBalance(address: string): Promise<number | string> {
  const url = "https://api.asteroidprotocol.io/v1/graphql";
  try {
    const query: Query = {
      query: `query {
        token_holder(offset: 0, limit: 100, where: {address: {_eq: "${address}"}}) {
          token {
            ticker
            content_path
            max_supply
            circulating_supply
            decimals
            last_price_base
            transaction {
              hash
            }
          }
          amount
          date_updated
        }
      }`,
      variables: {},
    };
    const response = await postJson(url, query) as ResponseRoids;
    let roidsAmount: number | null = null;
    if (
      typeof response === "object" &&
      response.data &&
      response.data.token_holder
    ) {
      for (const holder of response.data.token_holder) {
        if (holder.token && holder.token.ticker === "ROIDS") {
          roidsAmount = holder.amount;
          break;
        }
      }
    }
    if (roidsAmount === null) {
      return "Error getting data";
    }
    return roidsAmount / 1000000;
  } catch (err) {
    return "Error getting data: " + err.status;
  }
}

async function getListBalance(address: string): Promise<number | string> {
  const url = "https://api.asteroidprotocol.io/v1/graphql";
  try {
    const query: Query = {
      query: `query{
        token_open_position(where: {token_id: {_eq: 1}, seller_address: {_eq: "${address}"}, is_cancelled: {_eq: false}, is_filled: {_eq: false}})  {
          id
          ppt
          amount
          total
        }
      }`,
      variables: {},
    };
    const response = await postJson(url, query) as ResponseList;
    let listAmount: number | null = null;
    if (
      typeof response === "object" &&
      response.data &&
      response.data.token_open_position
    ) {
      if (response.data.token_open_position.length) {
        for (const open_position of response.data.token_open_position) {
          if (open_position.amount) {
            listAmount += open_position.amount;
          }
        }
      }
      else {
        return 0
      }
    }
    if (listAmount === null) {
      return "Error getting data";
    }
    return listAmount / 1000000;
  } catch (err) {
    return "Error getting data: " + err.status;
  }
}

async function getAtomBalance(address: string) {
  const url = `https://cosmos-api.polkachu.com/cosmos/bank/v1beta1/balances/${address}`;
  try {
    const response = await getJson(url);
    let atomAmount: number | null = null;
    if (typeof response === "object" && response.balances) {
      for (const item of response.balances) {
        if (item.denom === "uatom") {
          atomAmount = Number(item.amount);
          break;
        }
      }
    }
    if (atomAmount === null) {
      return "Error: no uatom found";
    }
    return (atomAmount / 1000000).toFixed(4);
  } catch (err) {
    return "Error getting data: " + err.status;
  }
}

async function main(workbook: ExcelScript.Workbook) {
  let selectedSheet = workbook.getActiveWorksheet();
  let addressRange = selectedSheet.getRange("A2:A51");
  let addressValue = addressRange.getValues();
  let rowCount = addressRange.getRowCount();
  for (let i = 0; i < rowCount; i++) {
    if (addressValue[i][0] as string !== "") {
      let roidsAmount = await getRoidsBalance(addressValue[i][0] as string);
      let listAmount = await getListBalance(addressValue[i][0] as string);
      selectedSheet.getCell(i + 1, 2).setValue(roidsAmount);
      selectedSheet.getCell(i + 1, 3).setValue(listAmount);
      selectedSheet.getCell(i + 1, 4).setFormula(`=B${i + 2} - C${i + 2} - D${i + 2}`)
      let atomAmount = await getAtomBalance(addressValue[i][0] as string);
      selectedSheet.getCell(i + 1, 5).setValue(atomAmount);
    }
  }
}