package Utilities;


class ExcelCellFormatting {
	private String fontName;
	
	String getFontName() {
		return fontName;
	}
	
	void setFontName(String fontName) {
		this.fontName = fontName;
	}
	
	private short fontSize;
	
	short getFontSize() {
		return fontSize;
	}
	
	void setFontSize(short fontSize) {
		this.fontSize = fontSize;
	}
	
	private short backColorIndex;
	
    short getBackColorIndex() {
    	return backColorIndex;
    }
    
    void setBackColorIndex(short backColorIndex) throws Exception {
    	if(backColorIndex < 0x8 || backColorIndex > 0x40) {
			throw new Exception("Valid indexes for the Excel custom palette are from 0x8 to 0x40 (inclusive)!");
		}
    	
    	this.backColorIndex = backColorIndex;
    }
    
    private short foreColorIndex;
    
    short getForeColorIndex() {
    	return foreColorIndex;
    }
    
   
    void setForeColorIndex(short foreColorIndex) throws Exception {
    	if(foreColorIndex < 0x8 || foreColorIndex > 0x40) {
			throw new Exception("Valid indexes for the Excel custom palette are from 0x8 to 0x40 (inclusive)!");
		}
    	
    	this.foreColorIndex = foreColorIndex;
    }
	
   
    boolean bold = false;
    
    boolean italics = false;
    
    boolean centred = false;
}