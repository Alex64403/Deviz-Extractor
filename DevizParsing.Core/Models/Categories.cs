namespace DevizParsing.Core.Models
{
    /// <summary>
    /// Stochează defalcarea pe categorii pentru materiale, manoperă, utilaje și transport.
    /// </summary>
    public class Categories
    {
        public Category Materials { get; set; } = new Category();
        public Category Labor { get; set; } = new Category();
        public Category Equipment { get; set; } = new Category();
        public Category Transport { get; set; } = new Category();
    }
}
