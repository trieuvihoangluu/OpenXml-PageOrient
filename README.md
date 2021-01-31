# OpenXml-PageOrient
Insert page break and session break

/*
 * The way to inserting page break
 * 1. From beginning to first TFL => portrait
 * 2. First TFL => landscape
 * 3. Each TFL => 1 page
 *
 * To adapt the requirement
 * 1. 
 * - When see the 1st TFL, we set the previous <p> is portrait page => this will set all previous page to portrait
 * - Insert last render page break to current <p> => this will stop the setting go to next section
 * 2.
 * - Add the default page format is landscape, it will apply to all other pages
 * -- Insert/Edit last node of <body> to be a landscape format
 * 3. 
 * - When see second TFL and from that point to last TFL, insert last render page break and <br type=page>
 * - The first TFL is skipped this <br type=page>
 */
