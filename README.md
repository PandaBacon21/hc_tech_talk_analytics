# Healthcare Tech Talk Analytics

> ** Note **
> 
> This is a simple script to help automate the analysis of our quarterly tech talk presentations. This is not intended to be used for any other purposes other than to serve as an example. 


> The following sample application is a personal, open-source project shared by the app creator and not an officially supported Zoom Video Communications, Inc. sample application. Zoom Video Communications, Inc., its employees and affiliates are not responsible for the use and maintenance of this application. Please use this sample application for inspiration, exploration and experimentation at your own risk and enjoyment. You may reach out to the app creator and broader Zoom Developer community on https://devforum.zoom.us/ for technical discussion and assistance, but understand there is no service level agreement support for this application. Thank you and happy coding!

## Usage 
There is only one real dependency - Pandas. Pandas will install all of its specific dependencies. 

I prefer using Spyder for my analytics projects but prefer to avoid the bulk of Anaconda. If you have Spyder installed via Anaconda, that works fine.

If you prefer to use Spyder on it's own like I do, you will also need spyder and spyder-kernels. To simplify, you can just install everything from the ``` requirements.txt ``` file. 

When pulling the reports from your Zoom Event, clear the first rows of every file so that the column headers serve as the top row. 

See the ```get_dfs()``` function for the name requirements of the individual tabs of your xlxs (or csv) file. If using csv, just change the pandas function to read the csv properly.