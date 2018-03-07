package com.yexj.excelOperator;

import java.util.Objects;

public final class SheetUtils {

	private SheetUtils(){
		
	}
    /**
     * client should make sure clazz is not null and is annotated by Sheet
     */
    public static Sheet annotationObject(Class<?> clazz) {
        Objects.requireNonNull(clazz);

        Sheet anno = clazz.getAnnotation(Sheet.class);
        Objects.requireNonNull(anno);

        return anno;
    }
}
