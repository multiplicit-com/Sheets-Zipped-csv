// test.js on GitHub
console.log('Script loaded on page:', window.location.pathname);

const uniqueID = 'HARDCODED_UNIQUE_ID';

if (window.location.pathname.includes('/thank-you')) {
  alert('Script is working on the Thank You page! Your unique ID is: ' + uniqueID);
} else {
  console.log('This is not the Thank You page');
}
