/**
 * Google Apps Script to fetch BNB Smart Chain transactions for two addresses
 * and output to a Google Sheet using BscScan API.
 * Addresses are read from "Addresses" sheet, output written to "Transactions" sheet.
 * Requires BscScan API key.
 */

// Configuration
const API_KEY = "YOUR_BSCSCAN_API_KEY"; // Replace with your BscScan API key
const MAINNET_API = "https://api.bscscan.com/api";
const BSCSCAN_URL = "https://bscscan.com/tx/";
const ADDRESS_SHEET = "Addresses"; // Sheet name for addresses
const OUTPUT_SHEET = "Transactions"; // Sheet name for output
const ADDRESS_CELLS = ["A1", "A2"]; // Cells containing the two addresses

/**
 * Main function to fetch and process transactions
 */
function fetchBscTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const addressSheet = ss.getSheetByName(ADDRESS_SHEET);
  const outputSheet = ss.getSheetByName(OUTPUT_SHEET);

  if (!addressSheet || !outputSheet) {
    Logger.log("Error: 'Addresses' or 'Transactions' sheet not found.");
    return;
  }

  // Read addresses
  const addresses = ADDRESS_CELLS.map(cell => 
    (addressSheet.getRange(cell).getValue() || "").trim().toLowerCase()
  ).filter(addr => isValidAddress(addr));

  if (addresses.length === 0) {
    Logger.log("Error: No valid addresses found in specified cells.");
    return;
  }

  // Prepare output sheet
  const headers = [
    "Timestamp", "Transaction Type", "Token", "Amount", 
    "From Address", "To Address", "Transaction Hash", "BscScan Link"
  ];
  outputSheet.clear();
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  outputSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");

  // Fetch and process transactions for each address
  let allTransactions = [];
  addresses.forEach(address => {
    const normalTxs = fetchNormalTransactions(address);
    const tokenTxs = fetchTokenTransactions(address);
    allTransactions = allTransactions.concat(
      processNormalTransactions(normalTxs, address),
      processTokenTransactions(tokenTxs, address)
    );
  });

  // Sort transactions by timestamp (descending)
  allTransactions.sort((a, b) => b.timestamp - a.timestamp);

  // Write transactions to sheet
  if (allTransactions.length > 0) {
    const data = allTransactions.map(tx => [
      new Date(tx.timestamp * 1000), // Convert UNIX timestamp to Date
      tx.type,
      tx.token,
      tx.amount,
      tx.from,
      tx.to,
      tx.hash,
      tx.link
    ]);
    outputSheet.getRange(2, 1, data.length, headers.length).setValues(data);
    
    // Format timestamp column
    outputSheet.getRange(2, 1, data.length, 1).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  } else {
    Logger.log("No transactions found for the provided addresses.");
  }
}

/**
 * Validates a BNB Smart Chain address
 * @param {string} address - Address to validate
 * @returns {boolean} - True if valid, false otherwise
 */
function isValidAddress(address) {
  return /^0x[a-fA-F0-9]{40}$/.test(address);
}

/**
 * Fetches normal transactions (BNB transfers) for an address
 * @param {string} address - BNB Smart Chain address
 * @returns {Object[]} - Array of transaction objects
 */
function fetchNormalTransactions(address) {
  const url = `${MAINNET_API}?module=account&action=txlist&address=${address}&startblock=0&endblock=99999999&sort=desc&apikey=${API_KEY}`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    Utilities.sleep(250); // Respect API rate limit (5 calls/sec)
    const json = JSON.parse(response.getContentText());
    if (json.status === "1") {
      return json.result;
    } else {
      Logger.log(`Error fetching normal transactions for ${address}: ${json.message}`);
      return [];
    }
  } catch (e) {
    Logger.log(`Error fetching normal transactions for ${address}: ${e.message}`);
    return [];
  }
}

/**
 * Fetches BEP-20 token transfer events for an address
 * @param {string} address - BNB Smart Chain address
 * @returns {Object[]} - Array of token transfer events
 */
function fetchTokenTransactions(address) {
  const url = `${MAINNET_API}?module=account&action=tokentx&address=${address}&startblock=0&endblock=99999999&sort=desc&apikey=${API_KEY}`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    Utilities.sleep(250); // Respect API rate limit
    const json = JSON.parse(response.getContentText());
    if (json.status === "1") {
      return json.result;
    } else {
      Logger.log(`Error retrieving token transactions for ${address}: ${json.message}`);
      return [];
    }
  } catch (e) {
    Logger.log(`Error retrieving token transactions for ${address}: ${e.message}`);
    return [];
  }
}

/**
 * Processes normal transactions to extract relevant data
 * @param {Object[]} transactions - Array of normal transaction objects
 * @param {string} address - Reference address to determine in/out
 * @returns {Object[]} - Processed transaction data
 */
function processNormalTransactions(transactions, address) {
  return transactions.map(tx => {
    const isOut = tx.from.toLowerCase() === address;
    const valueBNB = Number(tx.value) / 1e18; // Convert wei to BNB
    return {
      timestamp: Number(tx.timeStamp),
      type: isOut ? "OUT" : "IN",
      token: "BNB",
      amount: valueBNB,
      from: tx.from.toLowerCase(),
      to: tx.to.toLowerCase(),
      hash: tx.hash,
      link: `${BSCSCAN_URL}${tx.hash}`
    };
  }).filter(tx => tx.amount > 0); // Exclude zero-value transactions
}

/**
 * Processes BEP-20 token transfer events to extract relevant data
 * @param {Object[]} transactions - Array of token transfer events
 * @param {string} address - Reference address to determine in/out
 * @returns {Object[]} - Processed transaction data
 */
function processTokenTransactions(transactions, address) {
  return transactions.map(tx => {
    const isOut = tx.from.toLowerCase() === address;
    const decimals = Number(tx.tokenDecimal) || 18;
    const amount = Number(tx.value) / Math.pow(10, decimals);
    return {
      timestamp: Number(tx.timeStamp),
      type: isOut ? "OUT" : "IN",
      token: tx.tokenSymbol || "Unknown",
      amount: amount,
      from: tx.from.toLowerCase(),
      to: tx.to.toLowerCase(),
      hash: tx.hash,
      link: `${BSCSCAN_URL}${tx.hash}`
    };
  }).filter(tx => tx.amount > 0); // Exclude zero-value transactions
}
