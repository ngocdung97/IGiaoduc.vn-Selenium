using System.Collections.Generic;

namespace IGiaoDucSelenium
{
    public class ObjValue
    {
        public string ID { get; set; }
        public string Title { get; set; }
        public List<Lesson> Lessons { get; set; }
    }
    public class Lesson
    {
        public string ID { get; set; }
        public string TitleID { get; set; }
        public string Content { get; set; }
        public List<Exercise> Exercises { get; set; }
    }

    public class Exercise
    {
        public string ID { get; set; }
        public string LessonID { get; set; }
        public string Content { get; set; }
        public string Link { get; set; }
        public string Frame { get; set; }

    }
}
