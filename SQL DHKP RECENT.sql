USE [dbPajak]
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'SP_DHKP_SPPT')
BEGIN
    DROP PROC SP_DHKP_SPPT
END
GO
CREATE proc [SP_DHKP_SPPT]
@xTahun nvarchar(4)
as
begin
DECLARE @xBayar float,--,@xBayar2 float,
 @xProp nVarchar(2), @xKab nVarchar(2), @xxKec nVarchar(3), @xxKel nVarchar(3),@xxBlok nVarchar(3), @xxUrut nVarchar(4), @xxJenis nVarchar(1)
--DECLARE @X1 nVarchar(2), @X2 nVarchar(2), @X3 nVarchar(3), @X4 nVarchar(3),@X5 nVarchar(3), @X6 nVarchar(4), @X7 nVarchar(1)
	DECLARE C_SPPT1 CURSOR FOR 
	SELECT KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,
			PBB_YG_HARUS_DIBAYAR_SPPT,THN_PAJAK_SPPT FROM SPPT_1 WHERE THN_PAJAK_SPPT =@xTahun
	OPEN C_SPPT1
	FETCH FROM C_SPPT1 INTO @xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis, @xBayar,@xTahun
	WHILE @@FETCH_STATUS = 0
	BEGIN
			SELECT * FROM TEMP_DHKP WHERE NOPQ= @xProp +'.'+ @xKab +'.'+ @xxKec +'.'+ @xxKel +'.'+ @xxBlok +'-'+ @xxUrut +'.'+ @xxJenis AND ROUND(@xBayar,0)<>ROUND(PBB_YG_HARUS_DIBAYAR_SPPT,0)
			INSERT INTO TEMP_DHKP(JUM_UBAH) VALUES (@xBayar)
	FETCH NEXT FROM C_SPPT1 INTO @xProp, @xKab, @xxKec,@xxKel, @xxBlok, @xxUrut, @xxJenis, @xBayar,@xTahun
	END
	CLOSE C_SPPT1
	DEALLOCATE C_SPPT1
IF @@ERROR<>0 
BEGIN
ROLLBACK TRAN
END
ELSE
BEGIN
COMMIT TRAN
END
END