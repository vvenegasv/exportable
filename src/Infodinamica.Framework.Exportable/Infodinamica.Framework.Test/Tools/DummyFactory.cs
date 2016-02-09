using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Infodinamica.Framework.Test.Tools
{
    internal static class DummyFactory
    {
        private static readonly List<string> Names = new List<string>() { "Viki Vong", "Brinda Bookout", "Janean Jonas", "Von Vankeuren", "Katelyn Kulik", "Verona Valdez", "Arletha Ashbaugh", "Bong Bellew", "Ilona Irving", "Refugia Radney", "Lili Lechner", "Carl Cowens", "Shanae Sleeper", "Enedina Etter", "Luke Lemmer", "Leota Legree", "Liliana Locicero", "Clementina Cline", "Reed Rivenburg", "Solange Schartz", "Myrtle Martin", "Meghan Mikkelsen", "Shari Santana", "Stacie Selman", "Linsey Lofgren", "Dudley Donley", "Oleta Oleson", "Zandra Zick", "Dedra Durfee", "Zulema Zubia", "Maribel Mclane", "Breana Brannigan", "Mitchel Millikin", "Gregoria Gladstone", "Cristen Currey", "Yajaira Yamada", "Eufemia Easterday", "Randall Red", "Salley Saiz", "Hisako Heber", "Hunter Halbert", "Shawnta Selig", "Tracey Tannenbaum", "Alyce Apel", "Mariam Mccullen", "Elenore Etherton", "Merri Mood", "Sherrill Sakamoto", "Zonia Zwick", "Lauran Laforge", "Hilario Hartin", "Monserrate Marton", "Karole Koprowski", "Otha Ocasio", "Shaneka Sweeney", "Rowena Rinaldo", "Carylon Chamlee", "Corinna Coan", "Eun Eastin", "Carissa Christopherso", "Jeanine Juckett", "Niesha Nevius", "Goldie Graig", "Yessenia Yoshioka", "Berenice Baney", "Caroline Casper", "Trent Takahashi", "Enrique Esposito", "Senaida Shumway", "Nola Neill", "Seema Schuyler", "Sol Smithers", "Lynsey Lary", "Felipe Furman", "Drew Dison", "Ying Yearby", "Arianne Attaway", "Sabina Sidle", "Terrilyn Tello", "Sherrie Simonetti", "Leontine Layton", "Manuel Marmolejo", "Zofia Zajicek", "Mable Mayle", "Catina Chavous", "Monica Minjares", "Anika Alford", "Vanna Vaccaro", "Yolonda Yamauchi", "Jenine Jacinto", "Minta Matarazzo", "Lesli Lucus", "Lezlie Lehner", "Cris Coverdale", "Marquis Marquardt", "Jasmine Jeanbaptiste", "Howard Hemstreet", "Helga Hardnett", "Erik Edmonson", "Jenice Josephson", "Cathryn Crosby", "Lidia Loughlin", "Amira Amburn", "Shay Sisto", "Augustine Atha", "Terresa Tarantino", "Sharonda Schebler", "Ciara Clifton", "Ines Izzo", "Janella Jacinto", "Alexa Ammann", "Janina Janis", "Brinda Bernstein", "Vicente Volker", "Toby Tucker", "Delmer Dunsmore", "Tijuana Traywick", "Elmo Ellingson", "Karisa Koogler", "Shannan Schmucker", "Gustavo Gillan", "Erna Evelyn", "Nickolas Neuendorf", "Jackeline Jenny", "Ok Ohlinger", "Jacquelyn Jandreau", "Casandra Correll", "Deandrea Doke", "Chia Crochet", "Marylin Maynard", "Roland Ratcliff", "Carina Cuffee", "Everett Exley", "Leesa Lora", "Kendra Ketter", "Jenny Jetter", "Lenard Lester", "Mayola Mitton", "Johna Jeske", "Jason Janco", "Floy Freeland", "Domenica Duggan", "Alena Ackerman", "Lynelle Larrimore", "Teena Taillon", "Billye Bowes", "Kati Kantner", "Tonie Taub", "August Addison", "Elizabet Eckhoff", };
        private static readonly  Random Random = new Random();

        public static DummyPerson CreateDummyPerson()
        {
            var person = new DummyPerson();
            person.BirthDate = RandomBirthDate();
            person.Name = RandomName();
            person.Age = DateTime.Now.Year - person.BirthDate.Year;
            person.IsAdult = person.Age > 18;

            return person;
        }

        public static DummyPersonWithAttributes CreateDummyPersonWithAttributes()
        {
            var person = new DummyPersonWithAttributes();
            person.BirthDate = RandomBirthDate();
            person.Name = RandomName();
            person.Age = DateTime.Now.Year - person.BirthDate.Year;
            person.IsAdult = person.Age > 18;

            return person;
        }

        public static DummyPersonWIthSomeAttributes CreateDummyPersonWIthSomeAttributes()
        {
            var person = new DummyPersonWIthSomeAttributes();
            person.BirthDate = RandomBirthDate();
            person.Name = RandomName();
            person.Age = DateTime.Now.Year - person.BirthDate.Year;
            person.IsAdult = person.Age > 18;

            return person;
        }

        private static string RandomName()
        {
            int maxValue = Names.Count-1;
            int index = Random.Next(0, maxValue);
            return Names[index];
        }

        private static DateTime RandomBirthDate()
        {
            int year = Random.Next(1970, DateTime.Now.Year - 5);
            int month = Random.Next(1, 12);
            int day = Random.Next(1, 28);

            return new DateTime(year, month, day);
        }

    }
}
