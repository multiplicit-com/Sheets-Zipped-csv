// thankyou-script.js
const uniqueID = 'HARDCODED_UNIQUE_ID';

if (window.location.pathname.includes('/thank_you')) {
  // Display a popup when the Thank You page is loaded
  alert('Thank you for your order! Your unique ID is: ' + uniqueID);
}
