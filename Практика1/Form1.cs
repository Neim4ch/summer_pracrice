using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;

namespace Практика1         //запихал это в гит
{                           
    public partial class Form1 : Form
    {
        
       string  replace_enter_to_spacebar(string a)
        {
            char[] arr = a.ToCharArray();
                for (int i = 0; i < arr.Length; ++i)
            {
                if (arr[i] == '\n')
                { arr[i] = ' '; }
            }

            string str = new string(arr);
            return str;
        }


        string penalt = "0.3333333";
        int nomer = 1;
        string[,] array;
        string vhodnVail = "";
       // bool suffleanswers = false;

        public Form1()
        {
            InitializeComponent();
        }
        int vibor = 0;
      
        string raschet_fraction_plus(int kolvo) // плюсовой рассчет оценки нужно уточнить что это такое
        {
            switch (kolvo)
            {
                case 1:
                    {
                        return "100";
                       
                    }
                case 2:
                    {
                        return "50";
                      
                    }
                case 3:
                    {
                        return "33.33333";
                        
                    }
                case 4 :
                    {
                        return "25";
                        
                    }
                case 5:
                    {
                        return "20";
                        
                    }
                case 6:
                    {
                        return "16.66667";
                        
                    }
                case 7:
                    {
                        return "14.28571";
                    }
                case 8:
                    {
                        return "12.5";
                    }
                case 9:
                    {
                        return "11.11111";
                    }
                case 10:
                    {
                        return "10";
                    }              
            }return "0";    
                    
        }
        string raschet_fraction_minus(int gran1, int gran2, int kolvo) // минусовой рачет оцеки тоже уточнить
        {
            int a = ((gran2 - gran1) - kolvo -1);
            switch (a)
            {
                case 1:
                    {
                        return "-100";

                    }
                case 2:
                    {
                        return "-50";

                    }
                case 3:
                    {
                        return "-33.33333";

                    }
                case 4:
                    {
                        return "-25";

                    }
                case 5:
                    {
                        return "-20";

                    }
                case 6:
                    {
                        return "-16.66667";

                    }
                case 7:
                    {
                        return "-14.28571";
                    }
                case 8:
                    {
                        return "-12.5";
                    }
                case 9:
                    {
                        return "-11.11111";
                    }
                case 10:
                    {
                        return "-10";
                    }
                case 0:
                    {
                        return "0";
                    }
            }
            return "0";
        }
        int proverk(int gran1)  // какая-то проверка узнать что это и переименовть чтобы стало ясно
        {
            char[] c = array[1, gran1].ToCharArray();
            try
            {
                for (int i = 0; ; ++i)
                {
                    if ((c[i] == '_') && (c[i + 1] == '_') && (c[i + 2] == '_'))
                    { return 0; }

                }
            }
            catch { return 1; }        
        }
        /// ////////////////////
        public delegate void InvokeDelegate();
        public delegate void InvokeDelegate1();
        
        void brbar2() /* перед функцие был static выдавало ошибку я убрал хз с чем это связано ниже так же 
                         вообще брбар функции отвечают за заполнение прогресс баров  */
            {
            //progressBar2.Value = progressBar2.Value + 20; ;
            progressBar1.Value++;
        }
        void brbar1()
        {
             progressBar2.Value++;
        }

        string setAnswerNumbering()
        {
            int value = comboBoxNumeration.SelectedIndex;

            switch(value)
            {
                case 0: return "abc";
                case 1: return "ABC";       //некорректное значение нужно на курсах протестить
                case 2: return "123";       //некорректное значение нужно на курсах протестить
                case 3: return "I II III";  //некорректное значение нужно на курсах протестить
                case 4: return "i ii iii";  //некорректное значение нужно на курсах протестить
                default: return "none";
            }
        }

        void gategory(string a, XmlDocument document)   /* функция, которая добавляет в начале хмл информацию о курсе нужна только для этого и фактически работает всегда
                                                           с одним и тем же полем массива строк в котором указано название лекции [0, 1]*/
        {
            XmlNode element = document.CreateElement("question");
            document.DocumentElement.AppendChild(element); // указываем родителя
            XmlAttribute attribute = document.CreateAttribute("type"); // создаём атрибут
            attribute.Value = "category";
            element.Attributes.Append(attribute);

            XmlNode subElement1 = document.CreateElement("category"); // даём имя
            element.AppendChild(subElement1); // и указываем кому принадлежит

            XmlNode subsubElement1 = document.CreateElement("text"); // даём имя
            subsubElement1.InnerText = "$course$/"+a; // и значение
            subElement1.AppendChild(subsubElement1); // и указываем кому принадлежит
        }
        void opredelen(XmlDocument document, string name, int gran1, int gran2) /*Функция, которая определяет тип вопроса собственно тут нужно доделывать еще 2 типа вопросов
                                                                                 gran1 - граница вопроса сверху, gran2 - граница вопроса снизу*/
        {
            if (gran1 == gran2)
            { return; ; }
            if(name.Contains("[[") && name.Contains("]]"))//gapselect nuzhno testit
            {
                gapselect(document, name, gran1, gran2);
                return;
            }
            if ((array[2, gran1 + 1] != "") && (array[2, gran1 + 1] != "1"))
            {
                matching(document, name, gran1, gran2);///////сопоставление (1 ко многим?)
            }
            else
            {
                for (int j = 1; gran1 + j != gran2; j++)
                {
                   
                        if (array[2, gran1 + j] == "1")
                        {
                            kolvo_otvetov++;
                        }
                }

                // условие для тру фолс старое 
                /*
                                 if (((gran2 - gran1 - 1) == 2) &&( 1 == kolvo_otvetov)&& // ОООЧЕНЬ СТРАННОЕ УСЛОВИЕ ТУТ НАДО ПОСМОТРЕТЬ ЧТО ВООБЩЕ ПРОИСХОДИТ
                    (((array[1, gran1 + 2] == "  верно") ||
                        (array[1, gran1 + 2] == " верно") || 
                        (array[1, gran1 + 2] == " верно ") ||                   
                        (array[1, gran1 + 2] == "верно ") ||
                        (array[1, gran1 + 2] == "верно  ") ||
                        (array[1, gran1 + 2]=="верно") || 
                        (array[1, gran1 + 2] == "неверно") ||
                        (array[1, gran1 + 2] == "  неверно") ||
                        (array[1, gran1 + 2] == " неверно") ||
                        (array[1, gran1 + 2] == " неверно ") || 
                        (array[1, gran1 + 2] == "неверно ") || 
                    (array[1, gran1 + 2] == "неверно  ")) || 
                    ((array[1, gran1 + 1] == "  верно") || 
                        (array[1, gran1 + 1] == " верно") || 
                        (array[1, gran1 + 1] == " верно ") || 
                        (array[1, gran1 + 1] == "верно ") || 
                        (array[1, gran1 + 1] == "верно  ") || 
                        (array[1, gran1 + 1] == "верно") ||
                        (array[1, gran1 + 1] == "неверно") || 
                        (array[1, gran1 + 1] == "  неверно") || 
                        (array[1, gran1 + 1] == " неверно") || 
                        (array[1, gran1 + 1] == " неверно ") || 
                        (array[1, gran1 + 1] == "неверно ") || 
                    (array[1, gran1 + 1] == "неверно  "))))//true false
                 */

                // NEW СТРАННОЕ УСЛОВИЕ ТУТ НАДО ПОСМОТРЕТЬ ЧТО ВООБЩЕ ПРОИСХОДИТ.  Ну допустим это правильно ?
                if (((gran2 - gran1 - 1) == 2) &&( 1 == kolvo_otvetov) && 
                    (((array[1, gran1 + 2] == "  верно") && (array[1, gran1 + 2] == "неверно  ")) ||  
                    ((array[1, gran1 + 1] == "  верно") && (array[1, gran1 + 1] == "неверно  "))))      //true false
                {
                    try
                    {
                        truefalse(document, name, gran1, gran2);//++++ верно/неверно вопрос
                        kolvo_otvetov = 0;
                    }
                    catch { MessageBox.Show("1"); /*document.Save(XML)*/; } // видимо отладка
                }
                else
                {
                    if((gran2 - gran1 - 1) == 1 && (kolvo_otvetov == 1))
                    {
                        numerical(document, name, gran1, gran2);
                        kolvo_otvetov = 0;
                    }
                    if (((gran2 - gran1 - 1) > kolvo_otvetov) && (kolvo_otvetov == 1))//multichoice,ответов 1
                    {
                        try
                        {
                            multichoice_one(document, name, gran1, gran2);//++
                            kolvo_otvetov = 0;
                        }
                        catch { MessageBox.Show("2"); /*document.Save(XML) */; }
                    }

                    if (((gran2 - gran1 - 1) >= kolvo_otvetov) && (kolvo_otvetov > 1) && (proverk(gran1) == 1))//multichouice,ответов больше 1
                    {
                        try
                        {
                            multichoice(document, name, kolvo_otvetov, gran1, gran2);
                            kolvo_otvetov = 0;
                        }
                        catch { MessageBox.Show("3"); /*document.Save(XML) */; }
                    }
                    if ((gran2 - gran1 - 1) == kolvo_otvetov)//ввод названия
                    {
                        try
                        {
                            shortanswer(document, name, kolvo_otvetov, gran1, gran2);//++
                            kolvo_otvetov = 0;
                        }
                        catch { MessageBox.Show("4"); /*document.Save(XML) */; }
                    }
                }
            }

        }
       
        void multichoice_one(XmlDocument document, string nasv_vopr, int gran1, int gran2) // МНОЖЕСТВЕННЫЙ ВЫБОР (1) пропатчил шафл(протестить)
        {

            for (int i = gran1; i < gran2; ++i)
            {
                array[1, i] = replace_enter_to_spacebar(array[1, i]);
            }
            ////
            nasv_vopr = replace_enter_to_spacebar(nasv_vopr);
            ////
            XmlNode element = document.CreateElement("question");
            document.DocumentElement.AppendChild(element); // указываем родителя
            XmlAttribute attribute = document.CreateAttribute("type"); // создаём атрибут
            attribute.Value = "multichoice";
            element.Attributes.Append(attribute);

            XmlNode subElement1 = document.CreateElement("name"); // даём имя
            element.AppendChild(subElement1); // и указываем кому принадлежит

            XmlNode subsubElement1 = document.CreateElement("text"); // даём имя
            subsubElement1.InnerText = array[1, 1] + " " + array[0,gran1]; // и значение
            nomer++;
            subElement1.AppendChild(subsubElement1); // и указываем кому принадлежит
            /////решить проблему со знаками
            XmlNode questiontext = document.CreateElement("questiontext");
            element.AppendChild(questiontext); // указываем родителя
            XmlAttribute format = document.CreateAttribute("format"); // создаём атрибут
            format.Value = "html";
            questiontext.Attributes.Append(format);

            XmlNode subquestiontext = document.CreateElement("text"); // даём имя
            subquestiontext.InnerText =  nasv_vopr ; // и значение
            questiontext.AppendChild(subquestiontext); // и указываем кому принадлежит
            ////
            XmlNode generalfeedback = document.CreateElement("generalfeedback"); // даём имя
            element.AppendChild(generalfeedback); // и указываем кому принадлежит
            XmlAttribute format_generalfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_generalfeedback.Value = "html";
            generalfeedback.Attributes.Append(format_generalfeedback);

            XmlNode subgeneralfeedback = document.CreateElement("text"); // даём имя
            subgeneralfeedback.InnerText = ""; // и значение
            generalfeedback.AppendChild(subgeneralfeedback); // и указываем кому принадлежит
   
            XmlNode defaultgrade = document.CreateElement("defaultgrade");
            element.AppendChild(defaultgrade); // указываем родителя
            defaultgrade.InnerText = "1.0000000";
            
            XmlNode penalty = document.CreateElement("penalty");
            element.AppendChild(penalty); // указываем родителя
            penalty.InnerText = Convert.ToString(penalt);

            XmlNode hidden = document.CreateElement("hidden");
            element.AppendChild(hidden); // указываем родителя
            hidden.InnerText = "0";

            XmlNode single = document.CreateElement("single");
            element.AppendChild(single); // указываем родителя
            single.InnerText = "true";

            XmlNode shuffleanswers = document.CreateElement("shuffleanswers");
            element.AppendChild(shuffleanswers); // указываем родителя
            //shuffleanswers.InnerText = "true";
            shuffleanswers.InnerText = Convert.ToString(checkBoxShuffle.Checked);

            XmlNode answernumbering = document.CreateElement("answernumbering");
            element.AppendChild(answernumbering); // указываем родителя
            //answernumbering.InnerText = "abc";
            answernumbering.InnerText = setAnswerNumbering();

            XmlNode correctfeedback = document.CreateElement("correctfeedback"); // даём имя
            element.AppendChild(correctfeedback); // и указываем кому принадлежит
            XmlAttribute format_correctfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_correctfeedback.Value = "html";
            correctfeedback.Attributes.Append(format_correctfeedback);

            XmlNode subcorrectfeedback = document.CreateElement("text"); // даём имя
            subcorrectfeedback.InnerText = "Ваш ответ верный."; // и значение
            correctfeedback.AppendChild(subcorrectfeedback);  
            ///
            XmlNode partiallycorrectfeedback = document.CreateElement("partiallycorrectfeedback"); // даём имя
            element.AppendChild(partiallycorrectfeedback); // и указываем кому принадлежит
            XmlAttribute format_partiallycorrectfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_partiallycorrectfeedback.Value = "html";
            partiallycorrectfeedback.Attributes.Append(format_partiallycorrectfeedback);

            XmlNode subpartiallycorrectfeedback = document.CreateElement("text"); // даём имя
            subpartiallycorrectfeedback.InnerText = "Ваш ответ частично правильный."; // и значение
            partiallycorrectfeedback.AppendChild(subpartiallycorrectfeedback);
            ///
            XmlNode incorrectfeedback = document.CreateElement("incorrectfeedback"); // даём имя
            element.AppendChild(incorrectfeedback); // и указываем кому принадлежит
            XmlAttribute format_incorrectfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_incorrectfeedback.Value = "html";
            incorrectfeedback.Attributes.Append(format_incorrectfeedback);

            XmlNode subincorrectfeedback = document.CreateElement("text"); // даём имя
            subincorrectfeedback.InnerText = "Ваш ответ неправильный."; // и значение
            incorrectfeedback.AppendChild(subincorrectfeedback);
         
            XmlNode shownumcorrect = document.CreateElement("shownumcorrect");
            element.AppendChild(shownumcorrect); // указываем родителя
 
            for (int j = 0; j < (gran2 - gran1)-1;++j)
            {
                XmlNode answer = document.CreateElement("answer");
                element.AppendChild(answer); // указываем родителя
                XmlAttribute fract = document.CreateAttribute("fraction"); // создаём атрибут  
                if (array[2, gran1 + j + 1]  == "1")
                {
                    fract.Value = "100";
                }
                else { fract.Value = "0"; }
                answer.Attributes.Append(fract);

                XmlAttribute form = document.CreateAttribute("format");
                form.Value = "html";
                answer.Attributes.Append(form);
                //
                XmlNode subansw = document.CreateElement("text"); // даём имя          
                subansw.InnerText = array[1, gran1 + j+1]; // и значение
                answer.AppendChild(subansw);
                ////////
                XmlNode feedback = document.CreateElement("feedback"); // даём имя
                answer.AppendChild(feedback); // и указываем кому принадлежит
                XmlAttribute format_feedback = document.CreateAttribute("format"); // создаём атрибут
                format_feedback.Value = "html";
                feedback.Attributes.Append(format_feedback);

                XmlNode subfeedback = document.CreateElement("text"); // даём имя
                subfeedback.InnerText = ""; // и значение
                feedback.AppendChild(subfeedback); // и указываем кому принадлежит
            }
        }

        void multichoice(XmlDocument document,string nasv_vopr, int kolvo,int gran1 ,int gran2) // МНОЖЕСТВЕННЫЙ ВЫБОР (МНОГО) пропатчил шафл(протестить)
        {
            for (int i = gran1; i < gran2; ++i)
            {
                array[1, i] = replace_enter_to_spacebar(array[1, i]);
            }
            ////
            nasv_vopr = replace_enter_to_spacebar(nasv_vopr);
            XmlNode element = document.CreateElement("question");
            document.DocumentElement.AppendChild(element); // указываем родителя
            XmlAttribute attribute = document.CreateAttribute("type"); // создаём атрибут
            attribute.Value = "multichoice";
            element.Attributes.Append(attribute);

            XmlNode subElement1 = document.CreateElement("name"); // даём имя
            element.AppendChild(subElement1); // и указываем кому принадлежит

            XmlNode subsubElement1 = document.CreateElement("text"); // даём имя
            subsubElement1.InnerText = array[1,1]+" "+ array[0, gran1]; // и значение
            nomer++;
            subElement1.AppendChild(subsubElement1); // и указываем кому принадлежит
       
            XmlNode questiontext = document.CreateElement("questiontext");
            element.AppendChild(questiontext); // указываем родителя
            XmlAttribute format = document.CreateAttribute("format"); // создаём атрибут
            format.Value = "html";
            questiontext.Attributes.Append(format);

            XmlNode subquestiontext = document.CreateElement("text"); // даём имя
            subquestiontext.InnerText =  nasv_vopr ; // и значение
            questiontext.AppendChild(subquestiontext); // и указываем кому принадлежит
            ////
            XmlNode generalfeedback = document.CreateElement("generalfeedback"); // даём имя
            element.AppendChild(generalfeedback); // и указываем кому принадлежит
            XmlAttribute format_generalfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_generalfeedback.Value = "html";
            generalfeedback.Attributes.Append(format_generalfeedback);

            XmlNode subgeneralfeedback = document.CreateElement("text"); // даём имя
            subgeneralfeedback.InnerText =""; // и значение
            generalfeedback.AppendChild(subgeneralfeedback); // и указываем кому принадлежит
            ///////////////
            XmlNode defaultgrade = document.CreateElement("defaultgrade");
            element.AppendChild(defaultgrade); // указываем родителя
            defaultgrade.InnerText = "1.0000000";

            XmlNode penalty = document.CreateElement("penalty");
            element.AppendChild(penalty); // указываем родителя
            penalty.InnerText = Convert.ToString(penalt);    

            XmlNode hidden = document.CreateElement("hidden");
            element.AppendChild(hidden); // указываем родителя
            hidden.InnerText = "0";

            XmlNode single = document.CreateElement("single");
            element.AppendChild(single); // указываем родителя
            single.InnerText = "false";

            XmlNode shuffleanswers = document.CreateElement("shuffleanswers");
            element.AppendChild(shuffleanswers); // указываем родителя
            //shuffleanswers.InnerText = "true";
            shuffleanswers.InnerText = Convert.ToString(checkBoxShuffle.Checked);

            XmlNode answernumbering = document.CreateElement("answernumbering");
            element.AppendChild(answernumbering); // указываем родителя
            //answernumbering.InnerText = "abc";
            answernumbering.InnerText = setAnswerNumbering();


            XmlNode correctfeedback = document.CreateElement("correctfeedback"); // даём имя
            element.AppendChild(correctfeedback); // и указываем кому принадлежит
            XmlAttribute format_correctfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_correctfeedback.Value = "html";
            correctfeedback.Attributes.Append(format_correctfeedback);

            XmlNode subcorrectfeedback = document.CreateElement("text"); // даём имя
            subcorrectfeedback.InnerText = "Ваш ответ верный."; // и значение
            correctfeedback.AppendChild(subcorrectfeedback);
            ///
            XmlNode partiallycorrectfeedback = document.CreateElement("partiallycorrectfeedback"); // даём имя
            element.AppendChild(partiallycorrectfeedback); // и указываем кому принадлежит
            XmlAttribute format_partiallycorrectfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_partiallycorrectfeedback.Value = "html";
            partiallycorrectfeedback.Attributes.Append(format_partiallycorrectfeedback);

            XmlNode subpartiallycorrectfeedback = document.CreateElement("text"); // даём имя
            subpartiallycorrectfeedback.InnerText = "Ваш ответ частично правильный."; // и значение
            partiallycorrectfeedback.AppendChild(subpartiallycorrectfeedback);
            ///
            XmlNode incorrectfeedback = document.CreateElement("incorrectfeedback"); // даём имя
            element.AppendChild(incorrectfeedback); // и указываем кому принадлежит
            XmlAttribute format_incorrectfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_incorrectfeedback.Value = "html";
            incorrectfeedback.Attributes.Append(format_incorrectfeedback);

            XmlNode subincorrectfeedback = document.CreateElement("text"); // даём имя
            subincorrectfeedback.InnerText = "Ваш ответ неправильный."; // и значение
            incorrectfeedback.AppendChild(subincorrectfeedback);
            /////////////////////
            XmlNode shownumcorrect = document.CreateElement("shownumcorrect");
            element.AppendChild(shownumcorrect); // указываем родителя
                                  
            string fraction_plus = Convert.ToString(raschet_fraction_plus(kolvo));
            string fraction_minus = Convert.ToString(raschet_fraction_minus(gran1,gran2,kolvo));
            for (int j = 0; j < (gran2 - gran1)-1; ++j)
            {
                XmlNode answer = document.CreateElement("answer");
                element.AppendChild(answer); // указываем родителя
                XmlAttribute fract = document.CreateAttribute("fraction"); // создаём атрибут  
                if (array[2, gran1 + j + 1] == "1")
                {
                    fract.Value =fraction_plus;
                }
                else { fract.Value = fraction_minus; }
                answer.Attributes.Append(fract);

                XmlAttribute form = document.CreateAttribute("format");
                form.Value = "html";
                answer.Attributes.Append(form);
                //
                XmlNode subansw = document.CreateElement("text"); // даём имя          
                subansw.InnerText = array[1, gran1 + j + 1]; // и значение
                answer.AppendChild(subansw);
                ////////
                XmlNode feedback = document.CreateElement("feedback"); // даём имя
                answer.AppendChild(feedback); // и указываем кому принадлежит
                XmlAttribute format_feedback = document.CreateAttribute("format"); // создаём атрибут
                format_feedback.Value = "html";
                feedback.Attributes.Append(format_feedback);

                XmlNode subfeedback = document.CreateElement("text"); // даём имя
                subfeedback.InnerText = ""; // и значение
                feedback.AppendChild(subfeedback); // и указываем кому принадлежит
            }
       }

        void truefalse(XmlDocument document, string nasv_vopr, int gran1, int gran2) // ВЕРНО,НЕВЕРНО ТИП ВОПРОСА нет шафла
        {
            for (int i = gran1; i < gran2; ++i)
            {
                array[1, i] = replace_enter_to_spacebar(array[1, i]);
            }
            ////
            nasv_vopr = replace_enter_to_spacebar(nasv_vopr);
            XmlNode element = document.CreateElement("question");
            document.DocumentElement.AppendChild(element); // указываем родителя
            XmlAttribute attribute = document.CreateAttribute("type"); // создаём атрибут
            attribute.Value = "truefalse";
            element.Attributes.Append(attribute);

            XmlNode subElement1 = document.CreateElement("name"); // даём имя
            element.AppendChild(subElement1); // и указываем кому принадлежит

            XmlNode subsubElement1 = document.CreateElement("text"); // даём имя
            subsubElement1.InnerText = array[1, 1] + " " + array[0, gran1]; // и значение
            nomer++;
            subElement1.AppendChild(subsubElement1); // и указываем кому принадлежит
        ////
            XmlNode questiontext = document.CreateElement("questiontext");
            element.AppendChild(questiontext); // указываем родителя
            XmlAttribute format = document.CreateAttribute("format"); // создаём атрибут
            format.Value = "html";
            questiontext.Attributes.Append(format);

            XmlNode subquestiontext = document.CreateElement("text"); // даём имя
            subquestiontext.InnerText =   nasv_vopr ; // и значение
            questiontext.AppendChild(subquestiontext); // и указываем кому принадлежит
           ////
            XmlNode generalfeedback = document.CreateElement("generalfeedback"); // даём имя
            element.AppendChild(generalfeedback); // и указываем кому принадлежит
            XmlAttribute format_generalfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_generalfeedback.Value = "html";
            generalfeedback.Attributes.Append(format_generalfeedback);

            XmlNode subgeneralfeedback = document.CreateElement("text"); // даём имя
            subgeneralfeedback.InnerText = ""; // и значение
            generalfeedback.AppendChild(subgeneralfeedback); // и указываем кому принадлежит
            ////
            XmlNode defaultgrade = document.CreateElement("defaultgrade");
            element.AppendChild(defaultgrade); // указываем родителя
            defaultgrade.InnerText = "1.0000000";

            XmlNode penalty = document.CreateElement("penalty");
            element.AppendChild(penalty); // указываем родителя
            penalty.InnerText = "1.0000000";

            XmlNode hidden = document.CreateElement("hidden");
            element.AppendChild(hidden); // указываем родителя
            hidden.InnerText = "0";  
            for (int i = 1; i < 3; ++i)
            {
                XmlNode answer = document.CreateElement("answer");
                element.AppendChild(answer); // указываем родителя
                XmlAttribute fract = document.CreateAttribute("fraction"); // создаём атрибут
                if (array[2, gran1 + i] == "1")
                {
                    fract.Value = "100";
                }
                else
                {
                    fract.Value = "0";
                }
                answer.Attributes.Append(fract);

                XmlAttribute form = document.CreateAttribute("format");
                form.Value = "moodle_auto_format";
                answer.Attributes.Append(form);
                //
                XmlNode subansw = document.CreateElement("text"); // даём имя
                if (((array[1, gran1 + i]) == "верно")||((array[1, gran1 + i]) == "верно ") || ((array[1, gran1 + i]) == "верно  ") || ((array[1, gran1 + i]) == " верно ") || ((array[1, gran1 + i]) == "  верно") || ((array[1, gran1 + i]) == " верно"))
                {
                    subansw.InnerText = "true"; // и значение
                }
                else
                {
                    if ((array[1, gran1 + i] != "неверно")&& (array[1, gran1 + i] != "неверно ") && (array[1, gran1 + i] != "неверно  ") && (array[1, gran1 + i] != " неверно"))
                        { MessageBox.Show("Ошибка в вопросе номер [" + array[0, gran1] + "] типа ВЕРНО НЕВЕРНО - проверьте правильность возможного ответа(должно'верно' или 'неверно'                               у вас '"+ array[1, gran1 + i] + "' )"); }
                    subansw.InnerText = "false"; // и значени
                }
                answer.AppendChild(subansw);
              
                XmlNode feedback = document.CreateElement("feedback"); // даём имя
                answer.AppendChild(feedback); // и указываем кому принадлежит
                XmlAttribute format_feedback = document.CreateAttribute("format"); // создаём атрибут
                format_feedback.Value = "html";
                feedback.Attributes.Append(format_feedback);

                XmlNode subfeedback = document.CreateElement("text"); // даём имя
                subfeedback.InnerText = ""; // и значение
                feedback.AppendChild(subfeedback); // и указываем кому принадлежит
            }      
        }

        void shortanswer(XmlDocument document, string nasv_vopr, int kolvo, int gran1, int gran2) // КОРОТКИЙ ОТВЕТ (Я ТАК ПОНИМАЮ ВПИСАТЬ ПРОСТО) нет шафла
        {
            for (int i = gran1; i < gran2; ++i)
            {
                array[1, i] = replace_enter_to_spacebar(array[1, i]);
            }
            ////
            nasv_vopr = replace_enter_to_spacebar(nasv_vopr);
            XmlNode element = document.CreateElement("question");
            document.DocumentElement.AppendChild(element); // указываем родителя
            XmlAttribute attribute = document.CreateAttribute("type"); // создаём атрибут
            attribute.Value = "shortanswer";
            element.Attributes.Append(attribute);

            XmlNode subElement1 = document.CreateElement("name"); // даём имя
            element.AppendChild(subElement1); // и указываем кому принадлежит

            XmlNode subsubElement1 = document.CreateElement("text"); // даём имя
            subsubElement1.InnerText = array[1, 1] + " " + array[0, gran1]; // и значение
            nomer++;
            subElement1.AppendChild(subsubElement1); // и указываем кому принадлежит
                                                     ////
            XmlNode questiontext = document.CreateElement("questiontext");
            element.AppendChild(questiontext); // указываем родителя
            XmlAttribute format = document.CreateAttribute("format"); // создаём атрибут
            format.Value = "html";
            questiontext.Attributes.Append(format);

            XmlNode subquestiontext = document.CreateElement("text"); // даём имя
            subquestiontext.InnerText =  nasv_vopr ; // и значение
            questiontext.AppendChild(subquestiontext); // и указываем кому принадлежит
           
            XmlNode generalfeedback = document.CreateElement("generalfeedback"); // даём имя
            element.AppendChild(generalfeedback); // и указываем кому принадлежит
            XmlAttribute format_generalfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_generalfeedback.Value = "html";
            generalfeedback.Attributes.Append(format_generalfeedback);

            XmlNode subgeneralfeedback = document.CreateElement("text"); // даём имя
            subgeneralfeedback.InnerText = ""; // и значение
            generalfeedback.AppendChild(subgeneralfeedback); // и указываем кому принадлежит

            XmlNode defaultgrade = document.CreateElement("defaultgrade");
            element.AppendChild(defaultgrade); // указываем родителя
            defaultgrade.InnerText = "1.0000000";

            XmlNode penalty = document.CreateElement("penalty");
            element.AppendChild(penalty); // указываем родителя
            penalty.InnerText = Convert.ToString(penalt);

            XmlNode hidden = document.CreateElement("hidden");
            element.AppendChild(hidden); // указываем родителя
            hidden.InnerText = "0";

            XmlNode usecase = document.CreateElement("usecase");
            element.AppendChild(usecase); // указываем родителя
            usecase.InnerText = "0";         
            for (int i = 1; i < kolvo + 1; ++i)
            {
                XmlNode answer = document.CreateElement("answer");
                element.AppendChild(answer); // указываем родителя
                XmlAttribute fract = document.CreateAttribute("fraction"); // создаём атрибут        
                fract.Value = "100";
                answer.Attributes.Append(fract);

                XmlAttribute form = document.CreateAttribute("format");
                form.Value = "moodle_auto_format";
                answer.Attributes.Append(form);
                //
                XmlNode subansw = document.CreateElement("text"); // даём имя          
                subansw.InnerText = array[1, gran1 + i]; // и значение
                answer.AppendChild(subansw);
                ////////
                XmlNode feedback = document.CreateElement("feedback"); // даём имя
                answer.AppendChild(feedback); // и указываем кому принадлежит
                XmlAttribute format_feedback = document.CreateAttribute("format"); // создаём атрибут
                format_feedback.Value = "html";
                feedback.Attributes.Append(format_feedback);

                XmlNode subfeedback = document.CreateElement("text"); // даём имя
                subfeedback.InnerText = ""; // и значение
                feedback.AppendChild(subfeedback); // и указываем кому принадлежит
            }
        }

        void matching(XmlDocument document, string nasv_vopr, int gran1, int gran2) //СОПОСТАВЛЕНИЕ (1 КО МНОГИМ СУДЯ ПО ВСЕМУ) пропатчил шафл (тест)
        {
            for (int i = gran1; i < gran2; ++i)
            {
                array[1, i] = replace_enter_to_spacebar(array[1, i]);
            }
            ////
            nasv_vopr = replace_enter_to_spacebar(nasv_vopr);
            XmlNode element = document.CreateElement("question");
            document.DocumentElement.AppendChild(element); // указываем родителя
            XmlAttribute attribute = document.CreateAttribute("type"); // создаём атрибут
            attribute.Value = "matching";
            element.Attributes.Append(attribute);

            XmlNode subElement1 = document.CreateElement("name"); // даём имя
            element.AppendChild(subElement1); // и указываем кому принадлежит

            XmlNode subsubElement1 = document.CreateElement("text"); // даём имя
            subsubElement1.InnerText = array[1, 1] + " " + array[0, gran1]; // и значение
            nomer++;
            subElement1.AppendChild(subsubElement1); // и указываем кому принадлежит
       
            XmlNode questiontext = document.CreateElement("questiontext");
            element.AppendChild(questiontext); // указываем родителя
            XmlAttribute format = document.CreateAttribute("format"); // создаём атрибут
            format.Value = "html";
            questiontext.Attributes.Append(format);

            XmlNode subquestiontext = document.CreateElement("text"); // даём имя
            subquestiontext.InnerText = nasv_vopr ; // и значение
            questiontext.AppendChild(subquestiontext); // и указываем кому принадлежит

            XmlNode generalfeedback = document.CreateElement("generalfeedback"); // даём имя
            element.AppendChild(generalfeedback); // и указываем кому принадлежит
            XmlAttribute format_generalfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_generalfeedback.Value = "html";
            generalfeedback.Attributes.Append(format_generalfeedback);

            XmlNode subgeneralfeedback = document.CreateElement("text"); // даём имя
            subgeneralfeedback.InnerText = ""; // и значение
            generalfeedback.AppendChild(subgeneralfeedback); // и указываем кому принадлежит
     
            XmlNode defaultgrade = document.CreateElement("defaultgrade");
            element.AppendChild(defaultgrade); // указываем родителя
            defaultgrade.InnerText = "1.0000000";

            XmlNode penalty = document.CreateElement("penalty");
            element.AppendChild(penalty); // указываем родителя
            penalty.InnerText = Convert.ToString(penalt);

            XmlNode hidden = document.CreateElement("hidden");
            element.AppendChild(hidden); // указываем родителя
            hidden.InnerText = "0";
            XmlNode shuffleanswers = document.CreateElement("shuffleanswers");
            element.AppendChild(shuffleanswers); // указываем родителя
            //shuffleanswers.InnerText = "true";
            shuffleanswers.InnerText = Convert.ToString(checkBoxShuffle.Checked);

            XmlNode correctfeedback = document.CreateElement("correctfeedback"); // даём имя
            element.AppendChild(correctfeedback); // и указываем кому принадлежит
            XmlAttribute format_correctfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_correctfeedback.Value = "html";
            correctfeedback.Attributes.Append(format_correctfeedback);

            XmlNode subcorrectfeedback = document.CreateElement("text"); // даём имя
            subcorrectfeedback.InnerText = "Ваш ответ верный."; // и значение
            correctfeedback.AppendChild(subcorrectfeedback);
            ///
            XmlNode partiallycorrectfeedback = document.CreateElement("partiallycorrectfeedback"); // даём имя
            element.AppendChild(partiallycorrectfeedback); // и указываем кому принадлежит
            XmlAttribute format_partiallycorrectfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_partiallycorrectfeedback.Value = "html";
            partiallycorrectfeedback.Attributes.Append(format_partiallycorrectfeedback);

            XmlNode subpartiallycorrectfeedback = document.CreateElement("text"); // даём имя
            subpartiallycorrectfeedback.InnerText = "Ваш ответ частично правильный."; // и значение
            partiallycorrectfeedback.AppendChild(subpartiallycorrectfeedback);
            ///
            XmlNode incorrectfeedback = document.CreateElement("incorrectfeedback"); // даём имя
            element.AppendChild(incorrectfeedback); // и указываем кому принадлежит
            XmlAttribute format_incorrectfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_incorrectfeedback.Value = "html";
            incorrectfeedback.Attributes.Append(format_incorrectfeedback);

            XmlNode subincorrectfeedback = document.CreateElement("text"); // даём имя
            subincorrectfeedback.InnerText = "Ваш ответ неправильный."; // и значение
            incorrectfeedback.AppendChild(subincorrectfeedback);
      
            XmlNode shownumcorrect = document.CreateElement("shownumcorrect");
            element.AppendChild(shownumcorrect); // указываем родителя
           
            for (int i = 1;/* gran1+i != gran2*/; i++)
            {
                if ((array[1, gran1 + i] != "") && (array[2, gran1 + i] != ""))
                    {
                    XmlNode subquestion = document.CreateElement("subquestion"); // даём имя
                    element.AppendChild(subquestion); // и указываем кому принадлежит
                    XmlAttribute format_subquestion = document.CreateAttribute("format"); // создаём атрибут
                    format_subquestion.Value = "html";
                    subquestion.Attributes.Append(format_subquestion);

                    XmlNode _subquestion = document.CreateElement("text"); // даём имя
                    _subquestion.InnerText = array[1, gran1 + i]; // и значение
                    subquestion.AppendChild(_subquestion);

                    XmlNode answer = document.CreateElement("answer"); // даём имя
                    subquestion.AppendChild(answer); // и указываем кому принадлежит
                    XmlAttribute format_answer = document.CreateAttribute("format"); // создаём атрибут
                    format_answer.Value = "html";
                    subquestion.Attributes.Append(format_answer);

                    XmlNode _answer = document.CreateElement("text"); // даём имя
                    _answer.InnerText =  array[2, gran1 + i]  ; // и значение
                    answer.AppendChild(_answer);
                }

                if ((array[1, gran1 + i] == "") && (array[2, gran1 + i] != ""))
                {

                    XmlNode subquestion = document.CreateElement("subquestion"); // даём имя
                    element.AppendChild(subquestion); // и указываем кому принадлежит
                    XmlAttribute format_subquestion = document.CreateAttribute("format"); // создаём атрибут
                    format_subquestion.Value = "html";
                    subquestion.Attributes.Append(format_subquestion);

                    XmlNode _subquestion = document.CreateElement("text"); // даём имя
                    _subquestion.InnerText = ""; // и значение
                    subquestion.AppendChild(_subquestion);

                    XmlNode answer = document.CreateElement("answer"); // даём имя
                    subquestion.AppendChild(answer); // и указываем кому принадлежит
                    XmlAttribute format_answer = document.CreateAttribute("format"); // создаём атрибут
                    format_answer.Value = "html";
                    subquestion.Attributes.Append(format_answer);

                    XmlNode _answer = document.CreateElement("text"); // даём имя
                    _answer.InnerText =  array[2, gran1 + i] ; // и значение
                    answer.AppendChild(_answer);

                }

                if ((array[1, gran1 + i] == "") && (array[2, gran1 + i] == ""))
                {
                    break;
                }
            }

                }

        void numerical(XmlDocument document, string nasv_vopr, int gran1, int gran2)
        {
            for (int i = gran1; i < gran2; ++i)
            {
                array[1, i] = replace_enter_to_spacebar(array[1, i]);
            }
            ////
            nasv_vopr = replace_enter_to_spacebar(nasv_vopr);
            XmlNode element = document.CreateElement("question");
            document.DocumentElement.AppendChild(element); // указываем родителя
            XmlAttribute attribute = document.CreateAttribute("type"); // создаём атрибут
            attribute.Value = "numerical";
            element.Attributes.Append(attribute);

            XmlNode subElement1 = document.CreateElement("name"); // даём имя
            element.AppendChild(subElement1); // и указываем кому принадлежит

            XmlNode subsubElement1 = document.CreateElement("text"); // даём имя
            subsubElement1.InnerText = array[1, 1] + " " + array[0, gran1]; // и значение
            nomer++;
            subElement1.AppendChild(subsubElement1); // и указываем кому принадлежит

            XmlNode questiontext = document.CreateElement("questiontext");
            element.AppendChild(questiontext); // указываем родителя
            XmlAttribute format = document.CreateAttribute("format"); // создаём атрибут
            format.Value = "html";
            questiontext.Attributes.Append(format);

            XmlNode subquestiontext = document.CreateElement("text"); // даём имя
            subquestiontext.InnerText = nasv_vopr; // и значение
            questiontext.AppendChild(subquestiontext); // и указываем кому принадлежит
            ////
            XmlNode generalfeedback = document.CreateElement("generalfeedback"); // даём имя
            element.AppendChild(generalfeedback); // и указываем кому принадлежит
            XmlAttribute format_generalfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_generalfeedback.Value = "html";
            generalfeedback.Attributes.Append(format_generalfeedback);

            XmlNode subgeneralfeedback = document.CreateElement("text"); // даём имя
            subgeneralfeedback.InnerText = ""; // и значение
            generalfeedback.AppendChild(subgeneralfeedback); // и указываем кому принадлежит
            ///////////////
            XmlNode defaultgrade = document.CreateElement("defaultgrade");
            element.AppendChild(defaultgrade); // указываем родителя
            defaultgrade.InnerText = "1.0000000";

            XmlNode penalty = document.CreateElement("penalty");
            element.AppendChild(penalty); // указываем родителя
            penalty.InnerText = Convert.ToString(penalt);

            XmlNode hidden = document.CreateElement("hidden");
            element.AppendChild(hidden); // указываем родителя
            hidden.InnerText = "0";
            //////////////////////////////////////////
            
       //     XmlNode shuffleanswers = document.CreateElement("shuffleanswers");
       //     element.AppendChild(shuffleanswers); // указываем родителя
            //shuffleanswers.InnerText = "true";
       //     shuffleanswers.InnerText = Convert.ToString(checkBoxShuffle.Checked);

       //     string fraction_plus = Convert.ToString(raschet_fraction_plus(kolvo));
       //     string fraction_minus = Convert.ToString(raschet_fraction_minus(gran1, gran2, kolvo));
            XmlNode answer = document.CreateElement("answer");
            element.AppendChild(answer); // указываем родителя
            XmlAttribute fract = document.CreateAttribute("fraction"); // создаём атрибут  
         //       if (array[2, gran1 + j + 1] == "1")
         //       {
         //           fract.Value = fraction_plus;
         //       }
         //       else { fract.Value = fraction_minus; }
         //       answer.Attributes.Append(fract);

            fract.Value = "100";
            answer.Attributes.Append(fract);

            XmlAttribute answerformat = document.CreateAttribute("format");
            answerformat.Value = "moodle_auto_format";
            answer.Attributes.Append(answerformat);
           // answerformat.Value = "moodle_auto_format";
            //
            XmlNode subansw = document.CreateElement("text"); // даём имя          
            subansw.InnerText = array[1, gran1 + 1]; // и значение
            answer.AppendChild(subansw);
                ////////
            XmlNode feedback = document.CreateElement("feedback"); // даём имя
            answer.AppendChild(feedback); // и указываем кому принадлежит
            XmlAttribute format_feedback = document.CreateAttribute("format"); // создаём атрибут
            format_feedback.Value = "html";
            feedback.Attributes.Append(format_feedback);

            XmlNode subfeedback = document.CreateElement("text"); // даём имя
            subfeedback.InnerText = ""; // и значение
            feedback.AppendChild(subfeedback); // и указываем кому принадлежит

            XmlNode tolerance = document.CreateElement("tolerance");
            tolerance.InnerText = "0";
            answer.AppendChild(tolerance);

            XmlNode units = document.CreateElement("units");
            element.AppendChild(units);

            XmlNode subunits_unit = document.CreateElement("unit");
            units.AppendChild(subunits_unit);

            XmlNode subunit_multiplier = document.CreateElement("multiplier");
            subunits_unit.AppendChild(subunit_multiplier);
            subunit_multiplier.InnerText = "1";

            XmlNode subunit_unit_name = document.CreateElement("unit_name");/* здесь идет фича с забиванием ответов с единицами измерения
                                                                             * нужно чекнуть что да как когда сайт встанет */
            subunits_unit.AppendChild(subunit_unit_name);
            subunit_unit_name.InnerText = "";

            XmlNode unitgradingtype = document.CreateElement("unitgradingtype");/* разобраться с тем как 
                                                                                 * пишутся штрафы */
            element.AppendChild(unitgradingtype);
            unitgradingtype.InnerText = "0";    //?

            XmlNode unitpenalty = document.CreateElement("unitpenalty");
            element.AppendChild(unitpenalty);
            unitpenalty.InnerText = "0.1000000";    //?

            XmlNode showunits = document.CreateElement("showunits");
            element.AppendChild(showunits);
            showunits.InnerText = "0";

            XmlNode unitsleft = document.CreateElement("unitsleft");
            element.AppendChild(unitsleft);
            unitsleft.InnerText = "0";
        }

        void gapselect(XmlDocument document, string nasv_vopr, int gran1, int gran2)
        {
            for (int i = gran1; i < gran2; ++i)
            {
                array[1, i] = replace_enter_to_spacebar(array[1, i]);
            }
            ////
            nasv_vopr = replace_enter_to_spacebar(nasv_vopr);
            XmlNode element = document.CreateElement("question");
            document.DocumentElement.AppendChild(element); // указываем родителя
            XmlAttribute attribute = document.CreateAttribute("type"); // создаём атрибут
            attribute.Value = "gapselect";
            element.Attributes.Append(attribute);

            XmlNode subElement1 = document.CreateElement("name"); // даём имя
            element.AppendChild(subElement1); // и указываем кому принадлежит

            XmlNode subsubElement1 = document.CreateElement("text"); // даём имя
            subsubElement1.InnerText = array[1, 1] + " " + array[0, gran1]; // и значение
            nomer++;
            subElement1.AppendChild(subsubElement1); // и указываем кому принадлежит

            XmlNode questiontext = document.CreateElement("questiontext");
            element.AppendChild(questiontext); // указываем родителя
            XmlAttribute format = document.CreateAttribute("format"); // создаём атрибут
            format.Value = "html";
            questiontext.Attributes.Append(format);

            XmlNode subquestiontext = document.CreateElement("text"); // даём имя
            subquestiontext.InnerText = nasv_vopr; // и значение
            questiontext.AppendChild(subquestiontext); // и указываем кому принадлежит
            ////
            XmlNode generalfeedback = document.CreateElement("generalfeedback"); // даём имя
            element.AppendChild(generalfeedback); // и указываем кому принадлежит
            XmlAttribute format_generalfeedback = document.CreateAttribute("format"); // создаём атрибут
            format_generalfeedback.Value = "html";
            generalfeedback.Attributes.Append(format_generalfeedback);

            XmlNode subgeneralfeedback = document.CreateElement("text"); // даём имя
            subgeneralfeedback.InnerText = ""; // и значение
            generalfeedback.AppendChild(subgeneralfeedback); // и указываем кому принадлежит
            ///////////////
            XmlNode defaultgrade = document.CreateElement("defaultgrade");
            element.AppendChild(defaultgrade); // указываем родителя
            defaultgrade.InnerText = "1.0000000";

            XmlNode penalty = document.CreateElement("penalty");
            element.AppendChild(penalty); // указываем родителя
            penalty.InnerText = Convert.ToString(penalt);

            XmlNode hidden = document.CreateElement("hidden");
            element.AppendChild(hidden); // указываем родителя
            hidden.InnerText = "0";
            //////////////////////////////////////////

            XmlNode shuffleanswers = document.CreateElement("shuffleanswers");
            element.AppendChild(shuffleanswers);
            shuffleanswers.InnerText = "1";

            XmlNode correctfeedback = document.CreateElement("correctfeedback");
            element.AppendChild(correctfeedback);
            XmlAttribute format_correctfeedback = document.CreateAttribute("format");
            format_correctfeedback.Value = "html";
            correctfeedback.Attributes.Append(format_correctfeedback);

            XmlNode subcorrectfeedback = document.CreateElement("text");
            correctfeedback.AppendChild(subcorrectfeedback);
            subcorrectfeedback.InnerText = "";

            XmlNode partiallycorrectfeedback = document.CreateElement("partiallycorrectfeedback");
            element.AppendChild(partiallycorrectfeedback);
            XmlAttribute format_partiallycorrectfeedback = document.CreateAttribute("format");
            format_partiallycorrectfeedback.Value = "html";
            partiallycorrectfeedback.Attributes.Append(format_partiallycorrectfeedback);

            XmlNode subpartiallycorrectfeedback = document.CreateElement("text");
            partiallycorrectfeedback.AppendChild(subpartiallycorrectfeedback);
            subpartiallycorrectfeedback.InnerText = "";

            XmlNode incorrectfeedback = document.CreateElement("incorrectfeedback");
            element.AppendChild(incorrectfeedback);
            XmlAttribute format_incorrectfeedback = document.CreateAttribute("format");
            format_incorrectfeedback.Value = "html";
            incorrectfeedback.Attributes.Append(format_incorrectfeedback);

            XmlNode subincorrectfeedback = document.CreateElement("text");
            incorrectfeedback.AppendChild(subincorrectfeedback);
            subincorrectfeedback.InnerText = "";

            for (int j = 0; j < (gran2 - gran1) - 1; ++j)
            {
                XmlNode selectoption = document.CreateElement("selectoption");
                element.AppendChild(selectoption);

                XmlNode subSelectOptionText = document.CreateElement("text");
                selectoption.AppendChild(subSelectOptionText);
                subSelectOptionText.InnerText = array[1, gran1 + j + 1];

                XmlNode subSelectOptionGroup = document.CreateElement("group");
                selectoption.AppendChild(subSelectOptionGroup);
                subSelectOptionGroup.InnerText = array[2, gran1 + j + 1];
            }
        }

        public int rowCount;

        private void selectFileButton_Click(object sender, EventArgs e)  /*ВЫБОР ФАЙЛА(ТЕПЕРЬ ЭТО ОЧЕВИДНО)
                                                                          Выбираем файл пикаем в двумерный массив строк соответсвующий последней заполненной ячейке
                                                                          , то есть чуть ли не весь документ в целом претензий нет, однако возможно стоит ограничить область
                                                                           чтения. Хотя я подозреваю, что этот вопрос решается где-то в другом месте.*/
        {                                                                  
            ///////////////////////////Выбор файла и заполнение массива
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                vhodnVail = openFileDialog1.FileName;   // ПУТЬ К ФАЙЛУ
               textBox1.Text = openFileDialog1.FileName;
            }
            try
            {
                //Thread thr = new Thread(MyThreadFunction);
                   // progressBar2.Maximum = 80;
                    progressBar2.Value = 0;
                    progressBar1.Value = 0;
                Excel.Application ObjWorkExcel = new Excel.Application(); //открыть эксель
                                                                    //progressBar2.Invoke(new InvokeDelegate1(brbar2));////////////////////////////////
                Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(vhodnVail, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                                                                     //progressBar2.Invoke(new InvokeDelegate1(brbar2));//////////////////////////////
                Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1]; //получить 1 лист
                    //progressBar2.Invoke(new InvokeDelegate1(brbar2));///////////////////////
                var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);//1 ячейку
                                                                     //progressBar2.Invoke(new InvokeDelegate1(brbar2));////////////////////////////////////
                string[,] list = new string[lastCell.Column, lastCell.Row]; // массив значений с листа равен по размеру листу
                progressBar2.Maximum = /*lastCell.Column*/ 3 * lastCell.Row;
                rowCount = lastCell.Row;
                                                                     //progressBar1.Maximum = lastCell.Column * lastCell.Row;
                                                                     //progressBar1.Value = 0;
                for (int i = 0; i < 3/*lastCell.Column*/; i++) //по всем колонкам
                    for (int j = 0; j < lastCell.Row; j++) // по всем строкам
                    {
                        list[i, j] = ObjWorkSheet.Cells[j + 1, i + 1].Text.ToString();//считываем текст в строку

                        //progressBar1.Invoke(new InvokeDelegate(brbar1)); ;
                        progressBar2.Invoke(new InvokeDelegate(brbar1)); ;
                    }
                ObjWorkBook.Close(false, Type.Missing, Type.Missing); //закрыть не сохраняя
                array = list;
                ObjWorkExcel.Quit();
                vibor = 1;
            }
            catch { 
                MessageBox.Show("Возможно,вы выбрали файл неверного формата,пожалуйста,повторите выбор"); 
                progressBar2.Value = 0;
                //progressBar1.Value = 0;
            }
        }
       
        public int kolvo_otvetov = 0;
        

        private void generateXmlButton_Click(object sender, EventArgs e) //СГЕНЕРИТЬ ФАЙЛ
        {
            if (vibor == 0)
            {
                MessageBox.Show("Пожалуйста,сперва выберите файл");
            }
            else
            {
                string XML = textBox1.Text + ".xml";
                XmlTextWriter textWritter = new XmlTextWriter(XML, Encoding.UTF8);
                textWritter.WriteStartDocument();
                textWritter.WriteStartElement("quiz");
                textWritter.WriteEndElement();
                textWritter.Close();
                XmlDocument document = new XmlDocument();
                document.Load(XML);
                gategory(array[1, 0], document);
                int gran1 = 0;
                int gran2 = 0;
                int flag = 0;
                string name = "";

                progressBar1.Maximum = rowCount-1;
                progressBar1.Value = 0;

                for (int i = 3; ; i++)
                {
                    progressBar1.Invoke(new InvokeDelegate1(brbar2));
                    //MessageBox.Show(Convert.ToString(progressBar1.Value)) ;
                    try
                    {
                        if ((name == "") && (flag == 0))
                        {
                            flag = 1;
                            name = array[1, i];
                            gran1 = i;
                        }
                        if (array[1, i] == "")
                        {
                            gran2 = i;
                            opredelen(document, name, gran1, gran2);
                            name = "";
                            kolvo_otvetov = 0;
                            flag = 0;
                            gran1 = gran2 + 1;
                        }
                    }
                    catch
                    {
                        try
                        {
                            if (array[2, i] != "")
                            { gran1 = gran2 + 1; }
                            else
                            {
                                opredelen(document, name, gran1, i);
                                gran1 = gran2 + 1;
                            }
                        }
                        catch
                        {
                            try
                            {
                                string n = array[1, i - 1];
                                opredelen(document, name, gran1, i); gran1 = gran2 + 1; ;
                            }
                            catch
                            {
                                MessageBox.Show("Конвертирование завершено." + "\n" + "Итоговый файл и путь к нему(который совпадает с местоположением исходного *.xlsx) " + XML ); ; document.Save(XML); gran1 = 0; kolvo_otvetov = 0; gran2 = 0; flag = 0; name = ""; break;
                            }
                        }
                    }
                    document.Save(XML);
                    vibor = 0;
                }
            }
        }

        //private void label1_Click(object sender, EventArgs e) //*ВОПРОСИТЕЛЬНЫЙ ЗНАК*
        //{

        //}

    private void showAuthorsButton_Click(object sender, EventArgs e) // АВТОРЫ
        {
            Form2 frm2 = new Form2();
            frm2.Show();
        }

        private void button5_Click(object sender, EventArgs e) // ХЕЛП НОМЕР 1
        {
            Form3 frm3 = new Form3();
            frm3.Show();
        }

        private void button4_Click(object sender, EventArgs e) //ХЕЛП НОМЕР 2
        {
            Form4 frm4 = new Form4();
            frm4.Show();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
