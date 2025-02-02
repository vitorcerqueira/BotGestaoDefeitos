namespace BotGestaoDefeitos
{
    public  class Disciplina
    {
        public int linha { get; set; }
        public long? ID_REGISTRO { get; set; }
        public long? ID_DEFEITO { get; set; }
        public string ID_RONDA { get; set; }
        public string TIPO_INSPECAO { get; set; }
        public string DATA { get; set; }
        public string RESPONSAVEL { get; set; }
        public string STATUS { get; set; }
    }
}
