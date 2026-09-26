# Doctor Appointment Form

An online appointment booking form for a clinic. Patients fill in the form, and every booking is saved to an Excel sheet so the clinic staff can call them back.

## Features

- Clean, mobile-friendly form built with HTML and CSS
- Required fields and 10-digit phone number validation
- Confirmation page with a booking ID
- Bookings saved to `appointments.xlsx` with columns for staff to track calls (**Contacted?** dropdown and **Staff notes**)
- Bookings are not lost if the Excel file is open: they are added automatically once it is closed

## How to run

Requires Python 3.9 or newer.

```bash
pip install -r requirements.txt
python server.py
```

The form opens at http://localhost:8000. Press `Ctrl+C` in the terminal to stop the server.

## Project structure

- `index.html` - the appointment form
- `style.css` - styling
- `server.py` - serves the form and saves bookings to Excel

`appointments.xlsx` contains patient details, so it is excluded from Git.
