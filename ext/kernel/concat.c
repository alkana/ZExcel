
#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include <php.h>
#include "php_ext.h"
#include <ext/standard/php_string.h>
#include "ext.h"

#include "kernel/main.h"
#include "kernel/memory.h"
#include "kernel/concat.h"

void zephir_concat_sss(zval **result, const char *op1, zend_uint op1_len, const char *op2, zend_uint op2_len, const char *op3, zend_uint op3_len, int self_var TSRMLS_DC){

	zval result_copy;
	int use_copy = 0;
	uint offset = 0, length;

	length = op1_len + op2_len + op3_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, op2, op2_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len, op3, op3_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_sssssssssssssss(zval **result, const char *op1, zend_uint op1_len, const char *op2, zend_uint op2_len, const char *op3, zend_uint op3_len, const char *op4, zend_uint op4_len, const char *op5, zend_uint op5_len, const char *op6, zend_uint op6_len, const char *op7, zend_uint op7_len, const char *op8, zend_uint op8_len, const char *op9, zend_uint op9_len, const char *op10, zend_uint op10_len, const char *op11, zend_uint op11_len, const char *op12, zend_uint op12_len, const char *op13, zend_uint op13_len, const char *op14, zend_uint op14_len, const char *op15, zend_uint op15_len, int self_var TSRMLS_DC){

	zval result_copy;
	int use_copy = 0;
	uint offset = 0, length;

	length = op1_len + op2_len + op3_len + op4_len + op5_len + op6_len + op7_len + op8_len + op9_len + op10_len + op11_len + op12_len + op13_len + op14_len + op15_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, op2, op2_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len, op3, op3_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len, op4, op4_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len, op5, op5_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len, op6, op6_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len + op6_len, op7, op7_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len + op6_len + op7_len, op8, op8_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len + op6_len + op7_len + op8_len, op9, op9_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len + op6_len + op7_len + op8_len + op9_len, op10, op10_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len + op6_len + op7_len + op8_len + op9_len + op10_len, op11, op11_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len + op6_len + op7_len + op8_len + op9_len + op10_len + op11_len, op12, op12_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len + op6_len + op7_len + op8_len + op9_len + op10_len + op11_len + op12_len, op13, op13_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len + op6_len + op7_len + op8_len + op9_len + op10_len + op11_len + op12_len + op13_len, op14, op14_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + op5_len + op6_len + op7_len + op8_len + op9_len + op10_len + op11_len + op12_len + op13_len + op14_len, op15, op15_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_ssssvss(zval **result, const char *op1, zend_uint op1_len, const char *op2, zend_uint op2_len, const char *op3, zend_uint op3_len, const char *op4, zend_uint op4_len, zval *op5, const char *op6, zend_uint op6_len, const char *op7, zend_uint op7_len, int self_var TSRMLS_DC){

	zval result_copy, op5_copy;
	int use_copy = 0, use_copy5 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	length = op1_len + op2_len + op3_len + op4_len + Z_STRLEN_P(op5) + op6_len + op7_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, op2, op2_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len, op3, op3_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len, op4, op4_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len, Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + Z_STRLEN_P(op5), op6, op6_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + op2_len + op3_len + op4_len + Z_STRLEN_P(op5) + op6_len, op7, op7_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_sv(zval **result, const char *op1, zend_uint op1_len, zval *op2, int self_var TSRMLS_DC){

	zval result_copy, op2_copy;
	int use_copy = 0, use_copy2 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	length = op1_len + Z_STRLEN_P(op2);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_svs(zval **result, const char *op1, zend_uint op1_len, zval *op2, const char *op3, zend_uint op3_len, int self_var TSRMLS_DC){

	zval result_copy, op2_copy;
	int use_copy = 0, use_copy2 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	length = op1_len + Z_STRLEN_P(op2) + op3_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2), op3, op3_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_svsv(zval **result, const char *op1, zend_uint op1_len, zval *op2, const char *op3, zend_uint op3_len, zval *op4, int self_var TSRMLS_DC){

	zval result_copy, op2_copy, op4_copy;
	int use_copy = 0, use_copy2 = 0, use_copy4 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	length = op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2), op3, op3_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len, Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_svsvs(zval **result, const char *op1, zend_uint op1_len, zval *op2, const char *op3, zend_uint op3_len, zval *op4, const char *op5, zend_uint op5_len, int self_var TSRMLS_DC){

	zval result_copy, op2_copy, op4_copy;
	int use_copy = 0, use_copy2 = 0, use_copy4 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	length = op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4) + op5_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2), op3, op3_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len, Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4), op5, op5_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_svsvsv(zval **result, const char *op1, zend_uint op1_len, zval *op2, const char *op3, zend_uint op3_len, zval *op4, const char *op5, zend_uint op5_len, zval *op6, int self_var TSRMLS_DC){

	zval result_copy, op2_copy, op4_copy, op6_copy;
	int use_copy = 0, use_copy2 = 0, use_copy4 = 0, use_copy6 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op6) != IS_STRING) {
	   zend_make_printable_zval(op6, &op6_copy, &use_copy6);
	   if (use_copy6) {
	       op6 = &op6_copy;
	   }
	}

	length = op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4) + op5_len + Z_STRLEN_P(op6);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2), op3, op3_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len, Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4), op5, op5_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4) + op5_len, Z_STRVAL_P(op6), Z_STRLEN_P(op6));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy6) {
	   zval_dtor(op6);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_svsvsvs(zval **result, const char *op1, zend_uint op1_len, zval *op2, const char *op3, zend_uint op3_len, zval *op4, const char *op5, zend_uint op5_len, zval *op6, const char *op7, zend_uint op7_len, int self_var TSRMLS_DC){

	zval result_copy, op2_copy, op4_copy, op6_copy;
	int use_copy = 0, use_copy2 = 0, use_copy4 = 0, use_copy6 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op6) != IS_STRING) {
	   zend_make_printable_zval(op6, &op6_copy, &use_copy6);
	   if (use_copy6) {
	       op6 = &op6_copy;
	   }
	}

	length = op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4) + op5_len + Z_STRLEN_P(op6) + op7_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2), op3, op3_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len, Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4), op5, op5_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4) + op5_len, Z_STRVAL_P(op6), Z_STRLEN_P(op6));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4) + op5_len + Z_STRLEN_P(op6), op7, op7_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy6) {
	   zval_dtor(op6);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_svv(zval **result, const char *op1, zend_uint op1_len, zval *op2, zval *op3, int self_var TSRMLS_DC){

	zval result_copy, op2_copy, op3_copy;
	int use_copy = 0, use_copy2 = 0, use_copy3 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	length = op1_len + Z_STRLEN_P(op2) + Z_STRLEN_P(op3);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_svvvv(zval **result, const char *op1, zend_uint op1_len, zval *op2, zval *op3, zval *op4, zval *op5, int self_var TSRMLS_DC){

	zval result_copy, op2_copy, op3_copy, op4_copy, op5_copy;
	int use_copy = 0, use_copy2 = 0, use_copy3 = 0, use_copy4 = 0, use_copy5 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	length = op1_len + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, op1, op1_len);
	memcpy(Z_STRVAL_PP(result) + offset + op1_len, Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + op1_len + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4), Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vs(zval **result, zval *op1, const char *op2, zend_uint op2_len, int self_var TSRMLS_DC){

	zval result_copy, op1_copy;
	int use_copy = 0, use_copy1 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + op2_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), op2, op2_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vsv(zval **result, zval *op1, const char *op2, zend_uint op2_len, zval *op3, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op3_copy;
	int use_copy = 0, use_copy1 = 0, use_copy3 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), op2, op2_len);
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len, Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vsvs(zval **result, zval *op1, const char *op2, zend_uint op2_len, zval *op3, const char *op4, zend_uint op4_len, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op3_copy;
	int use_copy = 0, use_copy1 = 0, use_copy3 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3) + op4_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), op2, op2_len);
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len, Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3), op4, op4_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vsvsv(zval **result, zval *op1, const char *op2, zend_uint op2_len, zval *op3, const char *op4, zend_uint op4_len, zval *op5, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op3_copy, op5_copy;
	int use_copy = 0, use_copy1 = 0, use_copy3 = 0, use_copy5 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3) + op4_len + Z_STRLEN_P(op5);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), op2, op2_len);
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len, Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3), op4, op4_len);
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3) + op4_len, Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vsvsvs(zval **result, zval *op1, const char *op2, zend_uint op2_len, zval *op3, const char *op4, zend_uint op4_len, zval *op5, const char *op6, zend_uint op6_len, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op3_copy, op5_copy;
	int use_copy = 0, use_copy1 = 0, use_copy3 = 0, use_copy5 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3) + op4_len + Z_STRLEN_P(op5) + op6_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), op2, op2_len);
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len, Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3), op4, op4_len);
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3) + op4_len, Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3) + op4_len + Z_STRLEN_P(op5), op6, op6_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vsvv(zval **result, zval *op1, const char *op2, zend_uint op2_len, zval *op3, zval *op4, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op3_copy, op4_copy;
	int use_copy = 0, use_copy1 = 0, use_copy3 = 0, use_copy4 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3) + Z_STRLEN_P(op4);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), op2, op2_len);
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len, Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + op2_len + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vv(zval **result, zval *op1, zval *op2, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvs(zval **result, zval *op1, zval *op2, const char *op3, zend_uint op3_len, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + op3_len;
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), op3, op3_len);
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvsv(zval **result, zval *op1, zval *op2, const char *op3, zend_uint op3_len, zval *op4, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op4_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy4 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), op3, op3_len);
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + op3_len, Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvsvv(zval **result, zval *op1, zval *op2, const char *op3, zend_uint op3_len, zval *op4, zval *op5, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op4_copy, op5_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy4 = 0, use_copy5 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4) + Z_STRLEN_P(op5);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), op3, op3_len);
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + op3_len, Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + op3_len + Z_STRLEN_P(op4), Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvv(zval **result, zval *op1, zval *op2, zval *op3, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op3_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy3 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvvv(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op3_copy, op4_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy3 = 0, use_copy4 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvvvv(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op3_copy, op4_copy, op5_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy3 = 0, use_copy4 = 0, use_copy5 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4), Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvvvvvv(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, zval *op7, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op3_copy, op4_copy, op5_copy, op6_copy, op7_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy3 = 0, use_copy4 = 0, use_copy5 = 0, use_copy6 = 0, use_copy7 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	if (Z_TYPE_P(op6) != IS_STRING) {
	   zend_make_printable_zval(op6, &op6_copy, &use_copy6);
	   if (use_copy6) {
	       op6 = &op6_copy;
	   }
	}

	if (Z_TYPE_P(op7) != IS_STRING) {
	   zend_make_printable_zval(op7, &op7_copy, &use_copy7);
	   if (use_copy7) {
	       op7 = &op7_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4), Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5), Z_STRVAL_P(op6), Z_STRLEN_P(op6));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6), Z_STRVAL_P(op7), Z_STRLEN_P(op7));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy6) {
	   zval_dtor(op6);
	}

	if (use_copy7) {
	   zval_dtor(op7);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvvvvvvv(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, zval *op7, zval *op8, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op3_copy, op4_copy, op5_copy, op6_copy, op7_copy, op8_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy3 = 0, use_copy4 = 0, use_copy5 = 0, use_copy6 = 0, use_copy7 = 0, use_copy8 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	if (Z_TYPE_P(op6) != IS_STRING) {
	   zend_make_printable_zval(op6, &op6_copy, &use_copy6);
	   if (use_copy6) {
	       op6 = &op6_copy;
	   }
	}

	if (Z_TYPE_P(op7) != IS_STRING) {
	   zend_make_printable_zval(op7, &op7_copy, &use_copy7);
	   if (use_copy7) {
	       op7 = &op7_copy;
	   }
	}

	if (Z_TYPE_P(op8) != IS_STRING) {
	   zend_make_printable_zval(op8, &op8_copy, &use_copy8);
	   if (use_copy8) {
	       op8 = &op8_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4), Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5), Z_STRVAL_P(op6), Z_STRLEN_P(op6));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6), Z_STRVAL_P(op7), Z_STRLEN_P(op7));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7), Z_STRVAL_P(op8), Z_STRLEN_P(op8));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy6) {
	   zval_dtor(op6);
	}

	if (use_copy7) {
	   zval_dtor(op7);
	}

	if (use_copy8) {
	   zval_dtor(op8);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvvvvvvvv(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, zval *op7, zval *op8, zval *op9, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op3_copy, op4_copy, op5_copy, op6_copy, op7_copy, op8_copy, op9_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy3 = 0, use_copy4 = 0, use_copy5 = 0, use_copy6 = 0, use_copy7 = 0, use_copy8 = 0, use_copy9 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	if (Z_TYPE_P(op6) != IS_STRING) {
	   zend_make_printable_zval(op6, &op6_copy, &use_copy6);
	   if (use_copy6) {
	       op6 = &op6_copy;
	   }
	}

	if (Z_TYPE_P(op7) != IS_STRING) {
	   zend_make_printable_zval(op7, &op7_copy, &use_copy7);
	   if (use_copy7) {
	       op7 = &op7_copy;
	   }
	}

	if (Z_TYPE_P(op8) != IS_STRING) {
	   zend_make_printable_zval(op8, &op8_copy, &use_copy8);
	   if (use_copy8) {
	       op8 = &op8_copy;
	   }
	}

	if (Z_TYPE_P(op9) != IS_STRING) {
	   zend_make_printable_zval(op9, &op9_copy, &use_copy9);
	   if (use_copy9) {
	       op9 = &op9_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4), Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5), Z_STRVAL_P(op6), Z_STRLEN_P(op6));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6), Z_STRVAL_P(op7), Z_STRLEN_P(op7));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7), Z_STRVAL_P(op8), Z_STRLEN_P(op8));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8), Z_STRVAL_P(op9), Z_STRLEN_P(op9));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy6) {
	   zval_dtor(op6);
	}

	if (use_copy7) {
	   zval_dtor(op7);
	}

	if (use_copy8) {
	   zval_dtor(op8);
	}

	if (use_copy9) {
	   zval_dtor(op9);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvvvvvvvvv(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, zval *op7, zval *op8, zval *op9, zval *op10, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op3_copy, op4_copy, op5_copy, op6_copy, op7_copy, op8_copy, op9_copy, op10_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy3 = 0, use_copy4 = 0, use_copy5 = 0, use_copy6 = 0, use_copy7 = 0, use_copy8 = 0, use_copy9 = 0, use_copy10 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	if (Z_TYPE_P(op6) != IS_STRING) {
	   zend_make_printable_zval(op6, &op6_copy, &use_copy6);
	   if (use_copy6) {
	       op6 = &op6_copy;
	   }
	}

	if (Z_TYPE_P(op7) != IS_STRING) {
	   zend_make_printable_zval(op7, &op7_copy, &use_copy7);
	   if (use_copy7) {
	       op7 = &op7_copy;
	   }
	}

	if (Z_TYPE_P(op8) != IS_STRING) {
	   zend_make_printable_zval(op8, &op8_copy, &use_copy8);
	   if (use_copy8) {
	       op8 = &op8_copy;
	   }
	}

	if (Z_TYPE_P(op9) != IS_STRING) {
	   zend_make_printable_zval(op9, &op9_copy, &use_copy9);
	   if (use_copy9) {
	       op9 = &op9_copy;
	   }
	}

	if (Z_TYPE_P(op10) != IS_STRING) {
	   zend_make_printable_zval(op10, &op10_copy, &use_copy10);
	   if (use_copy10) {
	       op10 = &op10_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9) + Z_STRLEN_P(op10);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4), Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5), Z_STRVAL_P(op6), Z_STRLEN_P(op6));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6), Z_STRVAL_P(op7), Z_STRLEN_P(op7));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7), Z_STRVAL_P(op8), Z_STRLEN_P(op8));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8), Z_STRVAL_P(op9), Z_STRLEN_P(op9));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9), Z_STRVAL_P(op10), Z_STRLEN_P(op10));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy6) {
	   zval_dtor(op6);
	}

	if (use_copy7) {
	   zval_dtor(op7);
	}

	if (use_copy8) {
	   zval_dtor(op8);
	}

	if (use_copy9) {
	   zval_dtor(op9);
	}

	if (use_copy10) {
	   zval_dtor(op10);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvvvvvvvvvv(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, zval *op7, zval *op8, zval *op9, zval *op10, zval *op11, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op3_copy, op4_copy, op5_copy, op6_copy, op7_copy, op8_copy, op9_copy, op10_copy, op11_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy3 = 0, use_copy4 = 0, use_copy5 = 0, use_copy6 = 0, use_copy7 = 0, use_copy8 = 0, use_copy9 = 0, use_copy10 = 0, use_copy11 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	if (Z_TYPE_P(op6) != IS_STRING) {
	   zend_make_printable_zval(op6, &op6_copy, &use_copy6);
	   if (use_copy6) {
	       op6 = &op6_copy;
	   }
	}

	if (Z_TYPE_P(op7) != IS_STRING) {
	   zend_make_printable_zval(op7, &op7_copy, &use_copy7);
	   if (use_copy7) {
	       op7 = &op7_copy;
	   }
	}

	if (Z_TYPE_P(op8) != IS_STRING) {
	   zend_make_printable_zval(op8, &op8_copy, &use_copy8);
	   if (use_copy8) {
	       op8 = &op8_copy;
	   }
	}

	if (Z_TYPE_P(op9) != IS_STRING) {
	   zend_make_printable_zval(op9, &op9_copy, &use_copy9);
	   if (use_copy9) {
	       op9 = &op9_copy;
	   }
	}

	if (Z_TYPE_P(op10) != IS_STRING) {
	   zend_make_printable_zval(op10, &op10_copy, &use_copy10);
	   if (use_copy10) {
	       op10 = &op10_copy;
	   }
	}

	if (Z_TYPE_P(op11) != IS_STRING) {
	   zend_make_printable_zval(op11, &op11_copy, &use_copy11);
	   if (use_copy11) {
	       op11 = &op11_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9) + Z_STRLEN_P(op10) + Z_STRLEN_P(op11);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4), Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5), Z_STRVAL_P(op6), Z_STRLEN_P(op6));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6), Z_STRVAL_P(op7), Z_STRLEN_P(op7));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7), Z_STRVAL_P(op8), Z_STRLEN_P(op8));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8), Z_STRVAL_P(op9), Z_STRLEN_P(op9));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9), Z_STRVAL_P(op10), Z_STRLEN_P(op10));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9) + Z_STRLEN_P(op10), Z_STRVAL_P(op11), Z_STRLEN_P(op11));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy6) {
	   zval_dtor(op6);
	}

	if (use_copy7) {
	   zval_dtor(op7);
	}

	if (use_copy8) {
	   zval_dtor(op8);
	}

	if (use_copy9) {
	   zval_dtor(op9);
	}

	if (use_copy10) {
	   zval_dtor(op10);
	}

	if (use_copy11) {
	   zval_dtor(op11);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

void zephir_concat_vvvvvvvvvvvvvv(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, zval *op7, zval *op8, zval *op9, zval *op10, zval *op11, zval *op12, zval *op13, zval *op14, int self_var TSRMLS_DC){

	zval result_copy, op1_copy, op2_copy, op3_copy, op4_copy, op5_copy, op6_copy, op7_copy, op8_copy, op9_copy, op10_copy, op11_copy, op12_copy, op13_copy, op14_copy;
	int use_copy = 0, use_copy1 = 0, use_copy2 = 0, use_copy3 = 0, use_copy4 = 0, use_copy5 = 0, use_copy6 = 0, use_copy7 = 0, use_copy8 = 0, use_copy9 = 0, use_copy10 = 0, use_copy11 = 0, use_copy12 = 0, use_copy13 = 0, use_copy14 = 0;
	uint offset = 0, length;

	if (Z_TYPE_P(op1) != IS_STRING) {
	   zend_make_printable_zval(op1, &op1_copy, &use_copy1);
	   if (use_copy1) {
	       op1 = &op1_copy;
	   }
	}

	if (Z_TYPE_P(op2) != IS_STRING) {
	   zend_make_printable_zval(op2, &op2_copy, &use_copy2);
	   if (use_copy2) {
	       op2 = &op2_copy;
	   }
	}

	if (Z_TYPE_P(op3) != IS_STRING) {
	   zend_make_printable_zval(op3, &op3_copy, &use_copy3);
	   if (use_copy3) {
	       op3 = &op3_copy;
	   }
	}

	if (Z_TYPE_P(op4) != IS_STRING) {
	   zend_make_printable_zval(op4, &op4_copy, &use_copy4);
	   if (use_copy4) {
	       op4 = &op4_copy;
	   }
	}

	if (Z_TYPE_P(op5) != IS_STRING) {
	   zend_make_printable_zval(op5, &op5_copy, &use_copy5);
	   if (use_copy5) {
	       op5 = &op5_copy;
	   }
	}

	if (Z_TYPE_P(op6) != IS_STRING) {
	   zend_make_printable_zval(op6, &op6_copy, &use_copy6);
	   if (use_copy6) {
	       op6 = &op6_copy;
	   }
	}

	if (Z_TYPE_P(op7) != IS_STRING) {
	   zend_make_printable_zval(op7, &op7_copy, &use_copy7);
	   if (use_copy7) {
	       op7 = &op7_copy;
	   }
	}

	if (Z_TYPE_P(op8) != IS_STRING) {
	   zend_make_printable_zval(op8, &op8_copy, &use_copy8);
	   if (use_copy8) {
	       op8 = &op8_copy;
	   }
	}

	if (Z_TYPE_P(op9) != IS_STRING) {
	   zend_make_printable_zval(op9, &op9_copy, &use_copy9);
	   if (use_copy9) {
	       op9 = &op9_copy;
	   }
	}

	if (Z_TYPE_P(op10) != IS_STRING) {
	   zend_make_printable_zval(op10, &op10_copy, &use_copy10);
	   if (use_copy10) {
	       op10 = &op10_copy;
	   }
	}

	if (Z_TYPE_P(op11) != IS_STRING) {
	   zend_make_printable_zval(op11, &op11_copy, &use_copy11);
	   if (use_copy11) {
	       op11 = &op11_copy;
	   }
	}

	if (Z_TYPE_P(op12) != IS_STRING) {
	   zend_make_printable_zval(op12, &op12_copy, &use_copy12);
	   if (use_copy12) {
	       op12 = &op12_copy;
	   }
	}

	if (Z_TYPE_P(op13) != IS_STRING) {
	   zend_make_printable_zval(op13, &op13_copy, &use_copy13);
	   if (use_copy13) {
	       op13 = &op13_copy;
	   }
	}

	if (Z_TYPE_P(op14) != IS_STRING) {
	   zend_make_printable_zval(op14, &op14_copy, &use_copy14);
	   if (use_copy14) {
	       op14 = &op14_copy;
	   }
	}

	length = Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9) + Z_STRLEN_P(op10) + Z_STRLEN_P(op11) + Z_STRLEN_P(op12) + Z_STRLEN_P(op13) + Z_STRLEN_P(op14);
	if (self_var) {

		if (Z_TYPE_PP(result) != IS_STRING) {
			zend_make_printable_zval(*result, &result_copy, &use_copy);
			if (use_copy) {
				ZEPHIR_CPY_WRT_CTOR(*result, (&result_copy));
			}
		}

		offset = Z_STRLEN_PP(result);
		length += offset;
		Z_STRVAL_PP(result) = (char *) str_erealloc(Z_STRVAL_PP(result), length + 1);

	} else {
		Z_STRVAL_PP(result) = (char *) emalloc(length + 1);
	}

	memcpy(Z_STRVAL_PP(result) + offset, Z_STRVAL_P(op1), Z_STRLEN_P(op1));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1), Z_STRVAL_P(op2), Z_STRLEN_P(op2));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2), Z_STRVAL_P(op3), Z_STRLEN_P(op3));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3), Z_STRVAL_P(op4), Z_STRLEN_P(op4));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4), Z_STRVAL_P(op5), Z_STRLEN_P(op5));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5), Z_STRVAL_P(op6), Z_STRLEN_P(op6));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6), Z_STRVAL_P(op7), Z_STRLEN_P(op7));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7), Z_STRVAL_P(op8), Z_STRLEN_P(op8));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8), Z_STRVAL_P(op9), Z_STRLEN_P(op9));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9), Z_STRVAL_P(op10), Z_STRLEN_P(op10));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9) + Z_STRLEN_P(op10), Z_STRVAL_P(op11), Z_STRLEN_P(op11));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9) + Z_STRLEN_P(op10) + Z_STRLEN_P(op11), Z_STRVAL_P(op12), Z_STRLEN_P(op12));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9) + Z_STRLEN_P(op10) + Z_STRLEN_P(op11) + Z_STRLEN_P(op12), Z_STRVAL_P(op13), Z_STRLEN_P(op13));
	memcpy(Z_STRVAL_PP(result) + offset + Z_STRLEN_P(op1) + Z_STRLEN_P(op2) + Z_STRLEN_P(op3) + Z_STRLEN_P(op4) + Z_STRLEN_P(op5) + Z_STRLEN_P(op6) + Z_STRLEN_P(op7) + Z_STRLEN_P(op8) + Z_STRLEN_P(op9) + Z_STRLEN_P(op10) + Z_STRLEN_P(op11) + Z_STRLEN_P(op12) + Z_STRLEN_P(op13), Z_STRVAL_P(op14), Z_STRLEN_P(op14));
	Z_STRVAL_PP(result)[length] = 0;
	Z_TYPE_PP(result) = IS_STRING;
	Z_STRLEN_PP(result) = length;

	if (use_copy1) {
	   zval_dtor(op1);
	}

	if (use_copy2) {
	   zval_dtor(op2);
	}

	if (use_copy3) {
	   zval_dtor(op3);
	}

	if (use_copy4) {
	   zval_dtor(op4);
	}

	if (use_copy5) {
	   zval_dtor(op5);
	}

	if (use_copy6) {
	   zval_dtor(op6);
	}

	if (use_copy7) {
	   zval_dtor(op7);
	}

	if (use_copy8) {
	   zval_dtor(op8);
	}

	if (use_copy9) {
	   zval_dtor(op9);
	}

	if (use_copy10) {
	   zval_dtor(op10);
	}

	if (use_copy11) {
	   zval_dtor(op11);
	}

	if (use_copy12) {
	   zval_dtor(op12);
	}

	if (use_copy13) {
	   zval_dtor(op13);
	}

	if (use_copy14) {
	   zval_dtor(op14);
	}

	if (use_copy) {
	   zval_dtor(&result_copy);
	}

}

