package com.kaptan.reader;

import java.util.ArrayList;
import java.util.List;

public class ListOfStringData {

	private List<String> stringCellDatas;

	public ListOfStringData() {
		super();
	}

	public List<String> getStringCellDatas() {
		if (null == stringCellDatas) {
			this.stringCellDatas = new ArrayList<String>();
		}
		return stringCellDatas;
	}

	public void setStringCellDatas(List<String> stringCellDatas) {
		this.stringCellDatas = stringCellDatas;
	}

}
