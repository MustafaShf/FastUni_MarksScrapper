const puppeteer = require('puppeteer');
const XLSX = require('xlsx');

(async () => {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();

  // Navigate to the login page
  await page.goto('https://flexstudent.nu.edu.pk/Login', { timeout: 10000 });

  // Log in
  await page.type('#m_inputmask_4', '23L-3095');
  await page.type('#pass', '-');

  // Handle reCAPTCHA
  const frames = await page.frames();
  const recaptchaFrame = frames.find(frame => frame.url().includes('https://www.google.com/recaptcha/api2/anchor'));

  if (recaptchaFrame) {
    const checkbox = await recaptchaFrame.$('.recaptcha-checkbox');
    if (checkbox) {
      await checkbox.click();
    }
  }

  // Wait for manual captcha solving
  await new Promise(resolve => setTimeout(resolve, 30000));

  // Submit the login form
  await page.click('.m-form__actions #m_login_signin_submit');
  await page.waitForNavigation();

  // Wait for the sidebar to load and click it
  await page.waitForSelector('#m_aside_left', { visible: true });

  // Expand the sidebar menu if necessary
  const menuToggleButton = await page.$('#m_aside_left_offcanvas_toggle');
  if (menuToggleButton) {
    await menuToggleButton.click();
  }

  // Wait for 2 seconds after expanding the sidebar
  await new Promise(resolve => setTimeout(resolve, 2000));

  // Wait for the specific menu item
  const menuItems = await page.$$('.m-menu__nav > li');
  const marksMenuItemIndex = 3; // Adjust based on zero-based index
  if (menuItems[marksMenuItemIndex]) {
    const marksMenuItem = await menuItems[marksMenuItemIndex].$('a');
    if (marksMenuItem) {
      await marksMenuItem.click();
    }
  }

  await page.waitForNavigation();

  // Click on the dropdown to open it
  await page.click('#SemId');

  // Wait for the dropdown options to be visible
  await page.waitForSelector('#SemId option');
  
  // Select the second option (Spring 2024)
  const options = await page.$$('#SemId option');
  if (options.length > 1) {
    await options[1].evaluate(option => option.selected = true);
    await page.select('#SemId', '20241'); // Use the value of the second option
  }

  // Wait for the content to load or update after selection
  await new Promise(resolve => setTimeout(resolve, 2000));

  // Define subject IDs
  const subjects = [
    'CS1005',
    'EE1005',
    'EL1005',
    'MT1008',
    'SE1001'
  ];

  // Function to extract data from a card
  const extractCardData = async (cardId) => {
    return await page.evaluate((cardId) => {
      const cardElement = document.querySelector(`#${cardId} .card-body .sum_table`);
      if (cardElement) {
        const headers = Array.from(cardElement.querySelectorAll('thead th')).map(th => th.innerText);
        const rows = Array.from(cardElement.querySelectorAll('tbody tr')).map(tr => {
          const cells = Array.from(tr.querySelectorAll('td')).map(td => td.innerText);
          return cells;
        });
        return { headers, rows };
      }
      return null;
    }, cardId);
  };

  // Initialize array to store all the data
  const allData = {};

  // Extract data from each subject and store it
  for (const subject of subjects) {
    const subjectData = [];

    // Define card IDs to gather information from
    const cardIds = [
      `${subject}-Quiz`,
      `${subject}-Assignment`,
      `${subject}-Sessional-I`,
      `${subject}-Sessional-II`,
      `${subject}-Project`,
      `${subject}-Final_Exam`,
      `${subject}-Grand_Total_Marks`
    ];

    for (const cardId of cardIds) {
      const cardData = await extractCardData(cardId);
      if (cardData) {
        // Add card title
        subjectData.push([cardId]);

        // Add headers
        subjectData.push(cardData.headers);

        // Add rows
        for (const row of cardData.rows) {
          subjectData.push(row);
        }

        // Add an empty row for separation
        subjectData.push([]);
      }
    }

    allData[subject] = subjectData;
  }

  // Create a new workbook
  const workbook = XLSX.utils.book_new();

  // Add each subject's data to a new sheet
  for (const subject in allData) {
    const worksheet = XLSX.utils.aoa_to_sheet(allData[subject]);
    XLSX.utils.book_append_sheet(workbook, worksheet, subject);
  }

  // Write the workbook to a file
  XLSX.writeFile(workbook, 'MarksData.xlsx');

  await browser.close();
})();
