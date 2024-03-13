
let profile_data = [];
let completedRequests = 0;
let document_name = "";
const access_token = "your_access_token";


document.getElementById('fileInput').addEventListener('change', function (e) {
  var file = e.target.files[0];
  var reader = new FileReader();

  document_name = (document.getElementById('fileInput').files[0].name);
  reader.onload = function (e) {
    var data = new Uint8Array(e.target.result);
    var workbook = XLSX.read(data, { type: 'array', header: 1, raw: false, dateNF: 'dd/mm/yyyy;@', cellDates: true });
    var sheetName = workbook.SheetNames[0];
    var sheet = workbook.Sheets[sheetName];
    var jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, dateNF: 'dd/mm/yyyy;@', cellDates: true });

    printSearchResult(jsonData); // Display the parsed data in the console
  };

  reader.readAsArrayBuffer(file);
});

function insert_user_data(profile_data, response, user_data) {
  profile_data["#"] = response.excell_id;
  profile_data["Full name"] = user_data.fullname;
  profile_data["Birth date"] = `${String(user_data.birth_day).padStart(2, '0')}/${String(user_data.birth_month).padStart(2, '0')}/${user_data.birth_year}`;
  if (response.searchResponse != null) {
    profile_data["VK"] = "https://vk.com/id" + response.searchResponse.id;
    profile_data["VK profile status"] = response.searchResponse.is_closed ? 'Close profile' : 'Open profile';
    profile_data["VK subscriptions"] = response.subscriptions === "" ? 'No info' : response.subscriptions;
  } else {
    profile_data["VK"] = "No data";
    profile_data["VK profile status"] = "No data";
    profile_data["VK subscriptions"] = "No data";
  }
}

function writeIntoFile(profile_data) {
  // Create a new workbook and add a worksheet
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(profile_data);

  // Add the worksheet to the workbook
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");

  // Write the workbook to an Excel file
  XLSX.writeFile(wb, `Filled ${document_name}`);
}

function process_full_name(fullname) {
  let amount_of_words = (fullname.split(" ").length);

  if (amount_of_words >= 3) {
    for (let index = 2; index < amount_of_words; index++) {
      let lastIndex = fullname.lastIndexOf(" ");
      fullname = fullname.substring(0, lastIndex);
    }
  }
  return fullname;
}


async function executeMultipleMethods(user_data) {
  return new Promise((resolve, reject) => {
    const code = `
  var searchResponse;

  if (${user_data.birth_day} == 0) {
    searchResponse = API.users.search({
    q: "${user_data.fullname}",
    count: 1
    });

  } else {
    searchResponse = API.users.search({
    q: "${user_data.fullname}",
    birth_day: ${Number(user_data.birth_day)},
    birth_month: ${Number(user_data.birth_month)},
    birth_year: ${Number(user_data.birth_year)},
    count: 1
    });
  }

  var userId = searchResponse.items[0].id;

  var subscriptionsResponse = API.users.getSubscriptions({
    user_id: userId,
    count: 1
  });

  var subscriptionsResponsegroupscount = subscriptionsResponse.groups.count;
  var subscriptionsResponseuserscount = subscriptionsResponse.users.count;
  var allcount = subscriptionsResponsegroupscount + subscriptionsResponseuserscount;

  return {"searchResponse": searchResponse.items[0], "subscriptions": allcount};
`;

    const script = document.createElement("script");
    const url = `https://api.vk.com/method/execute?code=${encodeURIComponent(code)}&access_token=${access_token}&v=5.131&callback=my_callback`;

    script.src = url;
    window.my_callback = function (data) {
      resolve(data.response);
    };

    script.onerror = function (error) {
      reject(error);
    };

    document.getElementsByTagName("head")[0].appendChild(script);
  });
}

async function performLimitedRequest(user_data) {
  let response = null;

  // Make the request and handle rate limit
  while (!response) {
    response = await executeMultipleMethods(user_data);
    // Check if the response is null (indicating rate limit exceeded)
    if (!response) {
      console.log("Rate limit exceeded. Waiting before retrying...");
      await sleep(1000); // Wait for 1 second before retrying
    }
  }
  return response;
}

// Function to simulate asynchronous delay
function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

// Function to process the search result
async function printSearchResult(jsonData) {
  const totalRequests = jsonData.length - 1;

  for (let index = 1; index < jsonData.length; index++) {

    let user_data = {};
    user_data.fullname = process_full_name(jsonData[index][1]);



    if (jsonData[index][2] != undefined) {
      const [month, day, year] = jsonData[index][2].split('/').map(Number);
      user_data.birth_day = month;
      user_data.birth_month = day;
      user_data.birth_year = year;

    } else {
      user_data.birth_day = 0;
      user_data.birth_month = 0;
      user_data.birth_year = 0;
    }

    let response = await performLimitedRequest(user_data);
    response['excell_id'] = index;

    profile_data.push({});
    insert_user_data(profile_data[index - 1], response, user_data);

    document.getElementById("resuts-header").textContent = `Filled profiles ${completedRequests++ - 1} out of ${totalRequests}`;
  }

  document.getElementById("resuts-header").textContent = `VK profiles parsing completed`;
  writeIntoFile(profile_data);

  completedRequests = 0;
  profile_data = [];
}
