USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'xx_NJOPTKP')
BEGIN
    DROP PROC xx_NJOPTKP
END
GO
CREATE proc xx_NJOPTKP
@xTahun nvarchar(4)
AS
BEGIN
--CREATE TABLE #T_SEM(cID NVARCHAR(30),cKec nvarchar(3),cKel nvarchar(3),cBlok nvarchar(3),cUrut nvarchar(4),cJenis nvarchar(1),cNJOP bigint)
DECLARE @xID nvarchar (30),@xID1 nvarchar (30)
DECLARE @xxKec nVarchar(3)
DECLARE @xxKel nVarchar(3)
DECLARE @xxBlok nVarchar(3)
DECLARE @xxUrut nVarchar(4)
DECLARE @xxJenis nVarchar(1)
--DECLARE @xJBumi nVarchar(1)
DECLARE	@xNJOP bigint

DELETE FROM DAT_SUBJEK_PAJAK_NJOPTKP WHERE THN_NJOPTKP =@xTahun
SET NOCOUNT ON;

DECLARE C_BANGUNAN CURSOR FOR 
--SELECT SUBJEK_PAJAK_ID FROM QOBJEKPAJAK where JNS_BUMI='1' ORDER BY SUBJEK_PAJAK_ID ASC
SELECT SUBJEK_PAJAK_ID FROM QOBJEKPAJAK GROUP BY SUBJEK_PAJAK_ID, JNS_BUMI HAVING JNS_BUMI='1'
OPEN C_BANGUNAN
FETCH FROM C_BANGUNAN INTO @xID
WHILE @@FETCH_STATUS = 0
BEGIN	
	DECLARE C_BANGUNAN1 CURSOR FOR 
	SELECT SUBJEK_PAJAK_ID,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK, NO_URUT,KD_JNS_OP,NJOP_BUMI*1+NJOP_BNG*1 as tot_NJOP
	FROM QOBJEKPAJAK where SUBJEK_PAJAK_ID=@xID ORDER BY SUBJEK_PAJAK_ID ASC,NJOP_BUMI*1+NJOP_BNG*1 DESC
	OPEN C_BANGUNAN1
	FETCH FROM C_BANGUNAN1 INTO @xID1,@xxKec,@xxKel,@xxBlok,@xxURUT,@xxJenis,@xNJOP
	WHILE @@FETCH_STATUS = 0
	BEGIN
		--if @xID1 =@xID 
		--begin
	  		INSERT INTO DAT_SUBJEK_PAJAK_NJOPTKP (SUBJEK_PAJAK_ID, KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT, KD_JNS_OP,THN_NJOPTKP)
			VALUES (@xID1,'12','12',@xxKec,@xxKel,@xxBlok,@xxUrut,@xxJenis,@xTahun)
		--end
	FETCH NEXT FROM C_BANGUNAN1 INTO @xID1,@xxKec,@xxKel,@xxBlok,@xxURUT,@xxJenis,@xNJOP
	END
	CLOSE C_BANGUNAN1
	DEALLOCATE C_BANGUNAN1
	FETCH NEXT FROM C_BANGUNAN INTO @xID
END
CLOSE C_BANGUNAN
DEALLOCATE C_BANGUNAN
	


if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
END