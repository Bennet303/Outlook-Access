using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;


namespace ContactsTest
{
    [TestClass]
    public class ContactsTest
    {
        Outlook_Access.Contacts Contacts { get; set; }

        [TestInitialize]
        public void InitContacts()
        {
            Contacts = new Outlook_Access.Contacts(null, null, false, true);
        }

        [TestClass]
        public class TestFindContactsByFullName : ContactsTest
        {
            /// <summary>
            /// Tests wether the existing contact with the first name 'Test' and the 
            /// last name 'Contact' will be found by the function FindContactsByFullName
            /// </summary>
            [TestMethod]
            public void TestFindContactsByFullName_ExistingContact()
            {
                //Arrange
                string expectedFirstNane = "Test";
                string expectedLastName = "Contact";
                string actualFirstName = "";
                string actualLastName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByFullName(expectedFirstNane, expectedLastName);
                foreach (Outlook.ContactItem c in results)
                {
                    actualFirstName = c.FirstName;
                    actualLastName = c.LastName;
                }

                //Assert
                Assert.AreEqual(actualFirstName, expectedFirstNane, "The first name is not the first name of the found contact");
                Assert.AreEqual(actualFirstName, expectedFirstNane, "The last name is not the last name of the found contact");
            }

            /// <summary>
            /// Tests wether the non existing contact with the first name 'Test' and the 
            /// last name 'X' will be found by the function FindContactsByFullName
            /// </summary>
            [TestMethod]
            public void TestFindContactsByFullName_NonExistingContact()
            {
                //Arrange
                string searchedFirstNane = "Test";
                string searchedLastName = "X";
                string expectedLastName = "";
                string expectedFirstName = "";

                string actualFirstName = "";
                string actualLastName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByFullName(searchedFirstNane, searchedLastName);
                foreach (Outlook.ContactItem c in results)
                {
                    actualFirstName = c.FirstName;
                    actualLastName = c.LastName;
                }

                //Assert
                Assert.AreEqual(actualFirstName, expectedFirstName, "The first name is not the first name of the found contact");
                Assert.AreEqual(actualFirstName, expectedLastName, "The last name is not the last name of the found contact");
            }

        }

        [TestClass]
        public class TestFindContactsByFirstName : ContactsTest
        {   /// <summary>
            /// Tests wether the existing contact with the first name 'Test' will be found by 
            /// the function FindContactsByFirstName
            /// </summary>
            [TestMethod]
            public void TestFindContactsByFirstName_ExistingContact()
            {
                //Arrange
                string searchedFirstName = "Test";
                string expectedFirstName = "Test";

                string actualFirstName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByFirstName(searchedFirstName);
                foreach (Outlook.ContactItem c in results)
                {
                    actualFirstName = c.FirstName;
                }

                //Assert
                Assert.AreEqual(actualFirstName, expectedFirstName, "The first name is not the first name of the found contact");
            }

            /// <summary>
            /// Tests wether the non existing contact with the first name 'X' will be found by 
            /// the function FindContactsByFirstName
            /// </summary>
            [TestMethod]
            public void TestFindContactsByFirstName_NonExistingContact()
            {
                //Arrange
                string searchedFirstName = "X";
                string expectedFirstName = "";

                string actualFirstName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByFirstName(searchedFirstName);
                foreach (Outlook.ContactItem c in results)
                {
                    actualFirstName = c.FirstName;
                }

                //Assert
                Assert.AreEqual(actualFirstName, expectedFirstName, "No contact should have been found.");
            }

            /// <summary>
            /// Tests whether the existing contact with the first name 'Test' will be found by
            /// the function FindContactsByFirstName using the substring 'Tes'
            /// </summary>
            [TestMethod]
            public void TestFindContactsByFirstName_ExistingContact_BySubstring()
            {
                //Arrange
                string searchedFirstName = "Tes";
                string expectedFirstName = "Test";

                string actualFirstName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByFirstName(searchedFirstName, true);
                foreach (Outlook.ContactItem c in results)
                {
                    actualFirstName = c.FirstName;
                }

                //Assert
                Assert.AreEqual(actualFirstName, expectedFirstName, "The searched substring is not part of the first name of the found contact.");
            }

            /// <summary>
            /// Tests whether the non existing contact with the first name 'Uvwxyz' will be found by
            /// the function FindContactsByFirstName using the substring 'Xyz'
            /// </summary>
            [TestMethod]
            public void TestFindContactsByFirstName_NonExistingContact_BySubstring()
            {
                //Arrange
                string searchedFirstName = "Xyz";
                string expectedFirstName = "";

                string actualFirstName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByFirstName(searchedFirstName, true);
                foreach (Outlook.ContactItem c in results)
                {
                    actualFirstName = c.FirstName;
                }

                //Assert
                Assert.AreEqual(actualFirstName, expectedFirstName, "No contact should have been found.");
            }
        }

        [TestClass]
        public class TestFindContactsByLastName : ContactsTest
        {
            /// <summary>
            /// Tests wether the existing contact with the last name 'Contact' will be found by 
            /// the function FindContactsByLastName
            /// </summary>
            [TestMethod]
            public void TestFindContactsByLastName_ExistingContact()
            {
                //Arrange
                string searchedLastName = "Contact";
                string expectedLastName = "Contact";

                string actualLastName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByLastName(searchedLastName);
                foreach (Outlook.ContactItem c in results)
                {
                    actualLastName = c.LastName;
                }

                //Assert
                Assert.AreEqual(actualLastName, expectedLastName, "The last name is not the last name of the found contact");
            }

            /// <summary>
            /// Tests wether the non existing contact with the last name 'X' will be found by 
            /// the function FindContactsByLastName
            /// </summary>
            [TestMethod]
            public void TestFindContactsByLastName_NonExistingContact()
            {
                //Arrange
                string searchedLastName = "X";
                string expectedLastName = "";

                string actualLastName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByLastName(searchedLastName);
                foreach (Outlook.ContactItem c in results)
                {
                    actualLastName = c.LastName;
                }

                //Assert
                Assert.AreEqual(actualLastName, expectedLastName, "No contact should have been found");
            }

            /// <summary>
            /// Tests whether the existing contact with the last name 'Contact' will be found by
            /// the function FindContactsByLastName using the substring 'Conta'
            /// </summary>
            [TestMethod]
            public void TestFindContactsByLastName_ExistingContact_BySubstring()
            {
                //Arrange
                string searchedLastName = "Conta";
                string expectedLastName = "Contact";

                string actualLastName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByLastName(searchedLastName, true);
                foreach (Outlook.ContactItem c in results)
                {
                    actualLastName = c.LastName;
                }

                //Assert
                Assert.AreEqual(actualLastName, expectedLastName, "The searched substring is not part of the last name of the found contact.");
            }

            /// <summary>
            /// Tests whether a non existing contact will be found by
            /// the function FindContactsByLastName using the substring 'Xyz'
            /// </summary>
            [TestMethod]
            public void TestFindContactsByLastName_NonExistingContact_BySubstring()
            {
                //Arrange
                string searchedLastName = "Xyz";
                string expectedLastName = "";

                string actualLastName = "";

                //Act
                List<Outlook.ContactItem> results = Contacts.FindContactsByLastName(searchedLastName, true);
                foreach (Outlook.ContactItem c in results)
                {
                    actualLastName = c.LastName;
                }

                //Assert
                Assert.AreEqual(actualLastName, expectedLastName, "No contact should have been found.");
            }
        }
    }
}   