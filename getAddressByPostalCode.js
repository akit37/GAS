function getAddressByPostalCode(postalCode) {
  try {
    if (!/^\d{3}\-?\d{4}$/.test(postalCode)) throw new Error('Invalid postal code');
    var response = UrlFetchApp.fetch('http://zipcloud.ibsnet.co.jp/api/search?zipcode='+postalCode),
        result   = JSON.parse(response).results;
    if (response.getResponseCode() !== 200) {
      throw new Error('Unable to API');
    } else if(result == null) {
      throw new Error('Unavailable postal code');
    } else {
      return result[0].address1 + "," + result[0].address2 + "," + result[0].address3;
    }
  } catch(e) {
    return e.toString();
  }
}
