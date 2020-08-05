package org.ubfs.word.temp.annoation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface CovertEnum {
	

	Class<?> value(); 

	String methodName() default "getCodeByName";

}
