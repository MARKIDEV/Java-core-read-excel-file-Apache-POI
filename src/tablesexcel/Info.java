package tablesexcel;

public class Info {
	String label;
	String code;
	String format;
	String longueur;
	String valParDefaut;
	String utValParDef;
	String rubOblig;

	public Info(String label, String code, String format, String longueur,
			String valParDefaut, String utValParDef, String rubOblig) {
		this.label =  label ;
		this.code = code;
		this.format =  format;
		this.longueur =   longueur;
		this.valParDefaut = valParDefaut ;
		this.utValParDef = utValParDef ;
		this.rubOblig =  rubOblig.trim() ;

	}

	public Info() {
	}
	

	 @Override
	 public String toString() {
	 return label+ code + format
	 + longueur+ valParDefaut
	 + utValParDef+ rubOblig;
	
	 }
}
