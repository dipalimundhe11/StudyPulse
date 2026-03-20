let loginname; // Declare loginname outside the function
let signin = document.querySelector("#signin");

// add Event
signin.addEventListener("submit", function (event) {
    event.preventDefault();
    loginname = login(); // Assign the value returned by the login function to loginname
    let userinfo = JSON.parse(localStorage.getItem(loginname));
    if (loginname) {
        // Redirect to welcome.html with the username as a query parameter
        setTimeout(() => {
            location.href = 'welcome.html?username=' + encodeURIComponent(userinfo.flname);
        }, 2000);
    }
});

function login() {
    let emailInput = document.querySelector("#emailname");
    let passInput = document.querySelector("#passname");

    let loginname = emailInput.value.trim().toLowerCase();
    let loginpass = passInput.value;

    if (loginname === "" && loginpass === "") {
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'Please fill input fields'
        });
        return null; // Return null when the fields are not filled
    } else {
        let userinfo = JSON.parse(localStorage.getItem(loginname));
        if (userinfo !== null) {
            if (loginname == userinfo['emailname'] && loginpass == userinfo['passwordname']) {
                Swal.fire({
                    icon: 'success',
                    title: 'Good Job',
                    text: 'login sucessfullly',

                })
                
                // Record login in Excel
                ExcelHandler.recordLogin(userinfo['flname'], loginname);
                
                return loginname;
            }
            else {
                Swal.fire({
                    icon: 'error',
                    title: 'Oops...',
                    text: 'Password does not match',

                })

            }
            return null;
        }
        else {
            Swal.fire({
                icon: 'error',
                title: 'Oops...',
                text: 'Account does not exist',

            })
            return null;
        }
    }
}
