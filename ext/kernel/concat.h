
#ifndef ZEPHIR_KERNEL_CONCAT_H
#define ZEPHIR_KERNEL_CONCAT_H

#include <php.h>
#include <Zend/zend.h>

#include "kernel/main.h"

#define ZEPHIR_CONCAT_SSS(result, op1, op2, op3) \
	 zephir_concat_sss(&result, op1, sizeof(op1)-1, op2, sizeof(op2)-1, op3, sizeof(op3)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_SSS(result, op1, op2, op3) \
	 zephir_concat_sss(&result, op1, sizeof(op1)-1, op2, sizeof(op2)-1, op3, sizeof(op3)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_SV(result, op1, op2) \
	 zephir_concat_sv(&result, op1, sizeof(op1)-1, op2, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_SV(result, op1, op2) \
	 zephir_concat_sv(&result, op1, sizeof(op1)-1, op2, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_SVS(result, op1, op2, op3) \
	 zephir_concat_svs(&result, op1, sizeof(op1)-1, op2, op3, sizeof(op3)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_SVS(result, op1, op2, op3) \
	 zephir_concat_svs(&result, op1, sizeof(op1)-1, op2, op3, sizeof(op3)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_SVSVS(result, op1, op2, op3, op4, op5) \
	 zephir_concat_svsvs(&result, op1, sizeof(op1)-1, op2, op3, sizeof(op3)-1, op4, op5, sizeof(op5)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_SVSVS(result, op1, op2, op3, op4, op5) \
	 zephir_concat_svsvs(&result, op1, sizeof(op1)-1, op2, op3, sizeof(op3)-1, op4, op5, sizeof(op5)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_SVSVSVS(result, op1, op2, op3, op4, op5, op6, op7) \
	 zephir_concat_svsvsvs(&result, op1, sizeof(op1)-1, op2, op3, sizeof(op3)-1, op4, op5, sizeof(op5)-1, op6, op7, sizeof(op7)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_SVSVSVS(result, op1, op2, op3, op4, op5, op6, op7) \
	 zephir_concat_svsvsvs(&result, op1, sizeof(op1)-1, op2, op3, sizeof(op3)-1, op4, op5, sizeof(op5)-1, op6, op7, sizeof(op7)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_SVV(result, op1, op2, op3) \
	 zephir_concat_svv(&result, op1, sizeof(op1)-1, op2, op3, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_SVV(result, op1, op2, op3) \
	 zephir_concat_svv(&result, op1, sizeof(op1)-1, op2, op3, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VS(result, op1, op2) \
	 zephir_concat_vs(&result, op1, op2, sizeof(op2)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VS(result, op1, op2) \
	 zephir_concat_vs(&result, op1, op2, sizeof(op2)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VSV(result, op1, op2, op3) \
	 zephir_concat_vsv(&result, op1, op2, sizeof(op2)-1, op3, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VSV(result, op1, op2, op3) \
	 zephir_concat_vsv(&result, op1, op2, sizeof(op2)-1, op3, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VSVS(result, op1, op2, op3, op4) \
	 zephir_concat_vsvs(&result, op1, op2, sizeof(op2)-1, op3, op4, sizeof(op4)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VSVS(result, op1, op2, op3, op4) \
	 zephir_concat_vsvs(&result, op1, op2, sizeof(op2)-1, op3, op4, sizeof(op4)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VSVSV(result, op1, op2, op3, op4, op5) \
	 zephir_concat_vsvsv(&result, op1, op2, sizeof(op2)-1, op3, op4, sizeof(op4)-1, op5, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VSVSV(result, op1, op2, op3, op4, op5) \
	 zephir_concat_vsvsv(&result, op1, op2, sizeof(op2)-1, op3, op4, sizeof(op4)-1, op5, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VV(result, op1, op2) \
	 zephir_concat_vv(&result, op1, op2, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VV(result, op1, op2) \
	 zephir_concat_vv(&result, op1, op2, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VVS(result, op1, op2, op3) \
	 zephir_concat_vvs(&result, op1, op2, op3, sizeof(op3)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VVS(result, op1, op2, op3) \
	 zephir_concat_vvs(&result, op1, op2, op3, sizeof(op3)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VVSVV(result, op1, op2, op3, op4, op5) \
	 zephir_concat_vvsvv(&result, op1, op2, op3, sizeof(op3)-1, op4, op5, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VVSVV(result, op1, op2, op3, op4, op5) \
	 zephir_concat_vvsvv(&result, op1, op2, op3, sizeof(op3)-1, op4, op5, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VVV(result, op1, op2, op3) \
	 zephir_concat_vvv(&result, op1, op2, op3, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VVV(result, op1, op2, op3) \
	 zephir_concat_vvv(&result, op1, op2, op3, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VVVVS(result, op1, op2, op3, op4, op5) \
	 zephir_concat_vvvvs(&result, op1, op2, op3, op4, op5, sizeof(op5)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VVVVS(result, op1, op2, op3, op4, op5) \
	 zephir_concat_vvvvs(&result, op1, op2, op3, op4, op5, sizeof(op5)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VVVVVVS(result, op1, op2, op3, op4, op5, op6, op7) \
	 zephir_concat_vvvvvvs(&result, op1, op2, op3, op4, op5, op6, op7, sizeof(op7)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VVVVVVS(result, op1, op2, op3, op4, op5, op6, op7) \
	 zephir_concat_vvvvvvs(&result, op1, op2, op3, op4, op5, op6, op7, sizeof(op7)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VVVVVVVS(result, op1, op2, op3, op4, op5, op6, op7, op8) \
	 zephir_concat_vvvvvvvs(&result, op1, op2, op3, op4, op5, op6, op7, op8, sizeof(op8)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VVVVVVVS(result, op1, op2, op3, op4, op5, op6, op7, op8) \
	 zephir_concat_vvvvvvvs(&result, op1, op2, op3, op4, op5, op6, op7, op8, sizeof(op8)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VVVVVVVVS(result, op1, op2, op3, op4, op5, op6, op7, op8, op9) \
	 zephir_concat_vvvvvvvvs(&result, op1, op2, op3, op4, op5, op6, op7, op8, op9, sizeof(op9)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VVVVVVVVS(result, op1, op2, op3, op4, op5, op6, op7, op8, op9) \
	 zephir_concat_vvvvvvvvs(&result, op1, op2, op3, op4, op5, op6, op7, op8, op9, sizeof(op9)-1, 1 TSRMLS_CC);

#define ZEPHIR_CONCAT_VVVVVVVVVS(result, op1, op2, op3, op4, op5, op6, op7, op8, op9, op10) \
	 zephir_concat_vvvvvvvvvs(&result, op1, op2, op3, op4, op5, op6, op7, op8, op9, op10, sizeof(op10)-1, 0 TSRMLS_CC);
#define ZEPHIR_SCONCAT_VVVVVVVVVS(result, op1, op2, op3, op4, op5, op6, op7, op8, op9, op10) \
	 zephir_concat_vvvvvvvvvs(&result, op1, op2, op3, op4, op5, op6, op7, op8, op9, op10, sizeof(op10)-1, 1 TSRMLS_CC);


void zephir_concat_sss(zval **result, const char *op1, zend_uint op1_len, const char *op2, zend_uint op2_len, const char *op3, zend_uint op3_len, int self_var TSRMLS_DC);
void zephir_concat_sv(zval **result, const char *op1, zend_uint op1_len, zval *op2, int self_var TSRMLS_DC);
void zephir_concat_svs(zval **result, const char *op1, zend_uint op1_len, zval *op2, const char *op3, zend_uint op3_len, int self_var TSRMLS_DC);
void zephir_concat_svsvs(zval **result, const char *op1, zend_uint op1_len, zval *op2, const char *op3, zend_uint op3_len, zval *op4, const char *op5, zend_uint op5_len, int self_var TSRMLS_DC);
void zephir_concat_svsvsvs(zval **result, const char *op1, zend_uint op1_len, zval *op2, const char *op3, zend_uint op3_len, zval *op4, const char *op5, zend_uint op5_len, zval *op6, const char *op7, zend_uint op7_len, int self_var TSRMLS_DC);
void zephir_concat_svv(zval **result, const char *op1, zend_uint op1_len, zval *op2, zval *op3, int self_var TSRMLS_DC);
void zephir_concat_vs(zval **result, zval *op1, const char *op2, zend_uint op2_len, int self_var TSRMLS_DC);
void zephir_concat_vsv(zval **result, zval *op1, const char *op2, zend_uint op2_len, zval *op3, int self_var TSRMLS_DC);
void zephir_concat_vsvs(zval **result, zval *op1, const char *op2, zend_uint op2_len, zval *op3, const char *op4, zend_uint op4_len, int self_var TSRMLS_DC);
void zephir_concat_vsvsv(zval **result, zval *op1, const char *op2, zend_uint op2_len, zval *op3, const char *op4, zend_uint op4_len, zval *op5, int self_var TSRMLS_DC);
void zephir_concat_vv(zval **result, zval *op1, zval *op2, int self_var TSRMLS_DC);
void zephir_concat_vvs(zval **result, zval *op1, zval *op2, const char *op3, zend_uint op3_len, int self_var TSRMLS_DC);
void zephir_concat_vvsvv(zval **result, zval *op1, zval *op2, const char *op3, zend_uint op3_len, zval *op4, zval *op5, int self_var TSRMLS_DC);
void zephir_concat_vvv(zval **result, zval *op1, zval *op2, zval *op3, int self_var TSRMLS_DC);
void zephir_concat_vvvvs(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, const char *op5, zend_uint op5_len, int self_var TSRMLS_DC);
void zephir_concat_vvvvvvs(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, const char *op7, zend_uint op7_len, int self_var TSRMLS_DC);
void zephir_concat_vvvvvvvs(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, zval *op7, const char *op8, zend_uint op8_len, int self_var TSRMLS_DC);
void zephir_concat_vvvvvvvvs(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, zval *op7, zval *op8, const char *op9, zend_uint op9_len, int self_var TSRMLS_DC);
void zephir_concat_vvvvvvvvvs(zval **result, zval *op1, zval *op2, zval *op3, zval *op4, zval *op5, zval *op6, zval *op7, zval *op8, zval *op9, const char *op10, zend_uint op10_len, int self_var TSRMLS_DC);
#define zephir_concat_function(result, op1, op2) concat_function(result, op1, op2 TSRMLS_CC)

#endif /* ZEPHIR_KERNEL_CONCAT_H */

