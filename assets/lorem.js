const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];

const d = new Date();
let month = months[d.getMonth()];

const today = month + " " + d.getDate() + ", " + d.getFullYear();
const todayUpper = today.toUpperCase()
//console.log(today)

const fsHeader = ["Lorem Ipsum Inc.",
                  "1234 Dolor St.",
                  "Suite 100",
                  "Sit Amet, Consectetur 12345",
                  "",
                  "P 123-456-7890",
                  "F 098-765-4321"];

const fsIsPleased = ["Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
                     "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam.",
                     "Quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.",
                     "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."]

const sectionHeaders = ["Lorem", "Ipsum", "Dolor", "Sit", "Amet", "Consectetur", "Adipiscing", "Elit", "Sed", "Tempor"];

const documentStructure = {
    sections: [
        {
            title: "Lorem",
            beginText: "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
            list: {
                type: "labels",
                items: [
                    "Lorem 1",
                    "Lorem 2",
                    "Lorem 3",
                    "Lorem 4",
                    "Lorem 5"
                ],
                sublist: {
                    type: "number",
                    index: -1,
                    items: [
                        "Dolor 1",
                        "Dolor 2",
                        "Dolor 3",
                        "Dolor 4"
                    ]
                }
            },
            endText: "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
        },
        {
            title: "Ipsum",
            beginText: "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."
        },
        {
            title: "Dolor",
            list: {
                type: "letter",
                items: [
                    "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
                    "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
                    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat."
                ]
            }
        },
        {
            title: "Sit",
            beginText: "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur.",
            list: {
                type: "letter",
                items: [
                    "Excepteur sint occaecat cupidatat non proident.",
                    "Sunt in culpa qui officia deserunt mollit anim id est laborum."
                ]
            }
        },
        {
            title: "Amet",
            beginText: "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
            list: {
                type: "letter",
                items: [
                    "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
                    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.",
                    "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."
                ]
            }
        },
        {
            title: "Consectetur",
            beginText: "Excepteur sint occaecat cupidatat non proident.",
            list: {
                type: "letter",
                items: [
                    "Sunt in culpa qui officia deserunt mollit anim id est laborum.",
                    "Lorem ipsum dolor sit amet, consectetur adipiscing elit."
                ]
            }
        },
        {
            title: "Adipiscing",
            beginText: "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
            list: {
                type: "letter",
                items: [
                    "Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.",
                    "Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur."
                ]
            }
        },
        {
            title: "Elit",
            beginText: "Excepteur sint occaecat cupidatat non proident.",
            table: "table",
            endText: "Sunt in culpa qui officia deserunt mollit anim id est laborum."
        },
        {
            title: "Sed",
            beginText: "Lorem ipsum dolor sit amet, consectetur adipiscing elit.",
            list: {
                type: "hourly",
                items: [
                    "Principal $275/hour",
                    "Project Manager $195/hour",
                    "Project Coordinator $170/hour",
                    "Cost Estimating $155/hour",
                    "CAD $110/hour"
                ]
            }
        },
        {
            title: "Tempor",
            beginText: "Sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.",
            list: {
                type: "filloutForm",
                items: [
                    "Type or print name:",
                    "Signature:",
                    "Title:",
                    "Date:"
                ]
            }
        }
    ]
};

const data = {
    today: todayUpper,
    fsHeader: fsHeader,
    fsIsPleased: fsIsPleased,
    sectionHeaders: sectionHeaders,
    documentStructure: documentStructure
};

module.exports = data;
