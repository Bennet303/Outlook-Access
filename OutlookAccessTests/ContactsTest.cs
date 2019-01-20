using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;


namespace ContactsTest
{
    [TestClass]
    public class ContactsTest
    {
        [TestCleanup]
        public void CleanUpContacts()
        {
            //TODO: Remove Contact "Test" "Contact"
        }
        

        [TestClass]
        public class ContactsReadingTest : ContactsTest
        {
            Outlook_Access.Contacts Contacts { get; set; }

            [TestInitialize]
            public void InitContacts()
            {
                Contacts = new Outlook_Access.Contacts(null, null, false, true);

            }

            [TestClass]
            public class TestFindContactsByFullName : ContactsReadingTest
            {
                /// <summary>
                /// Tests wether the existing contact with the first name 'Test' and the 
                /// last name 'Contact' will be found by the function FindContactsByFullName
                /// </summary>
                [TestMethod]
                public void TestFindContactsByFullName_ExistingContact()
                {
                    //Arrange
                    const string searchedFirstName = "Test";
                    const string searchedLastName = "Contact";
                    string expectedFirstName = "Test";
                    string expectedLastName = "Contact";

                    string actualFirstName = "";
                    string actualLastName = "";

                    //Act
                    List<Outlook.ContactItem> results = Contacts.FindContactsByFullName(searchedFirstName, searchedLastName);
                    foreach (Outlook.ContactItem c in results)
                    {
                        actualFirstName = c.FirstName;
                        actualLastName = c.LastName;
                    }

                    //Assert
                    Assert.AreEqual(actualFirstName, expectedFirstName, "The first name is not the first name of the found contact");
                    Assert.AreEqual(actualLastName, expectedLastName, "The last name is not the last name of the found contact");
                }

                /// <summary>
                /// Tests wether the non existing contact with the first name 'Test' and the 
                /// last name 'X' will be found by the function FindContactsByFullName
                /// </summary>
                [TestMethod]
                public void TestFindContactsByFullName_NonExistingContact()
                {
                    //Arrange
                    const string searchedFirstName = "Test";
                    const string searchedLastName = "X";
                    const string expectedLastName = "";
                    const string expectedFirstName = "";

                    string actualFirstName = "";
                    string actualLastName = "";

                    //Act
                    List<Outlook.ContactItem> results = Contacts.FindContactsByFullName(searchedFirstName, searchedLastName);
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
            public class TestFindContactsByFirstName : ContactsReadingTest
            {   /// <summary>
                /// Tests wether the existing contact with the first name 'Test' will be found by 
                /// the function FindContactsByFirstName
                /// </summary>
                [TestMethod]
                public void TestFindContactsByFirstName_ExistingContact()
                {
                    //Arrange
                    const string searchedFirstName = "Test";
                    const string expectedFirstName = "Test";

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
                    const string searchedFirstName = "X";
                    const string expectedFirstName = "";

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
                    const string searchedFirstName = "Tes";
                    const string expectedFirstName = "Test";

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
                    const string searchedFirstName = "Xyz";
                    const string expectedFirstName = "";

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
            public class TestFindContactsByLastName : ContactsReadingTest
            {
                /// <summary>
                /// Tests wether the existing contact with the last name 'Contact' will be found by 
                /// the function FindContactsByLastName
                /// </summary>
                [TestMethod]
                public void TestFindContactsByLastName_ExistingContact()
                {
                    //Arrange
                    const string searchedLastName = "Contact";
                    const string expectedLastName = "Contact";

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
                    const string searchedLastName = "X";
                    const string expectedLastName = "";

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
                    const string searchedLastName = "onta";
                    const string expectedLastName = "Contact";

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
                    const string searchedLastName = "Xyz";
                    const string expectedLastName = "";

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

            [TestClass]
            public class TestFindContactsByEmail : ContactsReadingTest
            {
                [TestMethod]
                public void TestFindContactByEmail_ExistingContact()
                {
                    //Arrange
                    const string searchedEmail = "test@email.com";
                    const string expectedEmail = "test@email.com";

                    string actualEmail = "";

                    //Act
                    List<Outlook.ContactItem> results = Contacts.FindContactsByEmail(searchedEmail, false);
                    foreach (Outlook.ContactItem c in results)
                    {
                        actualEmail = c.Email1Address;
                    }

                    //Assert 
                    Assert.AreEqual(actualEmail, expectedEmail, "No contact with this email has been found.");
                }

                [TestMethod]
                public void TestFindContactByEmail_NonExistingContact()
                {
                    //Arrange
                    const string searchedEmail = "test@outlook.com";
                    const string expectedEmail = "";

                    string actualEmail = "";

                    //Act
                    List<Outlook.ContactItem> results = Contacts.FindContactsByEmail(searchedEmail, false);
                    foreach (Outlook.ContactItem c in results)
                    {
                        actualEmail = c.Email1Address;
                    }

                    //Assert 
                    Assert.AreEqual(actualEmail, expectedEmail, "No contact should have been found.");
                }

                [TestMethod]
                public void TestFindContactByEmail_ExistingContact_BySubstring()
                {
                    //Arrange
                    const string searchedEmail = "test@email.com";
                    const string expectedEmail = "test@email.com";

                    string actualEmail = "";

                    //Act
                    List<Outlook.ContactItem> results = Contacts.FindContactsByEmail(searchedEmail, false);
                    foreach (Outlook.ContactItem c in results)
                    {
                        actualEmail = c.Email1Address;
                    }

                    //Assert 
                    Assert.AreEqual(actualEmail, expectedEmail, "The searched substring is not part of the email of the found contact.");
                }

                [TestMethod]
                public void TestFindContactsByEmail_NonExistingContact_BySubstring()
                {
                    //Arrange
                    const string searchedEmail = "st@outlook.c";
                    const string expectedEmail = "";

                    string actualEmail = "";

                    //Act
                    List<Outlook.ContactItem> results = Contacts.FindContactsByEmail(searchedEmail, true);
                    foreach (Outlook.ContactItem c in results)
                    {
                        actualEmail = c.Email1Address;
                    }

                    //Assert 
                    Assert.AreEqual(actualEmail, expectedEmail, "No contact should have been found.");
                }
            }
        }
    }
}   