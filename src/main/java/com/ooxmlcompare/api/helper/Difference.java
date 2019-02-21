package com.ooxmlcompare.api.helper;

public enum Difference {
    DELETE_SHEET("deleteSheet"), 
    INSERT_SHEET("addSheet"),
    UPDATED_SHEET("updateSheet"),
	DELETE_CELL("deleteCell"),
	INSERT_CELL("insertCell"),
	UPDATE_CELL("updateCell");

    private String type;

    private Difference(String type) {
    	this.setType(type);
    }

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}
}
