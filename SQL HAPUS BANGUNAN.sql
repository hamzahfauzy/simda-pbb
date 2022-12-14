USE [dbPajak]
GO
/****** Object:  StoredProcedure [dbo].[HAPUS_BANGUNAN]    Script Date: 07/22/2014 12:42:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
ALTER proc [dbo].[HAPUS_BANGUNAN]
@xxKec nvarchar(3),@xxKel nvarchar(3),@xxBlok nvarchar(3),@xxUrut nvarchar(4),@xxJenis nvarchar(1),
@xNOP nvarchar(30),@xNo smallint,
@xNOP1 nvarchar(30),@xNo1 smallint,
@xNOP2 nvarchar(30),@xNo2 smallint,
--@xLuas bigint,@xNilai bigint,@xNOP3 nvarchar(30),
@xJPB nvarchar(2),
@xNOP4 nvarchar(30),@xNo4 smallint,
@xNOP5 nvarchar(30),@xNo5 smallint,
@xNOP6 nvarchar(30),@xNo6 smallint,
@xNOP7 nvarchar(30),@xNo7 smallint,
@xNOP8 nvarchar(30),@xNo8 smallint,
@xNOP9 nvarchar(30),@xNo9 smallint,
@xNOP10 nvarchar(30),@xNo10 smallint,
@xNOP11 nvarchar(30),@xNo11 smallint,
@xNOP12 nvarchar(30),@xNo12 smallint,
@xNOP13 nvarchar(30),@xNo13 smallint,
@xNOP14 nvarchar(30),@xNo14 smallint,
@xNOP15 nvarchar(30),@xNo15 smallint,
@xNOP16 nvarchar(30),@xNo16 smallint,
@xNOP17 nvarchar(30),@xNo17 smallint

as
begin
DECLARE @xLuas bigint
DECLARE @xNilai bigint

DELETE FROM DAT_OP_BANGUNAN where (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP) =  @xNOP  AND (NO_BNG  =@xNo)
DELETE FROM DAT_NILAI_INDIVIDU where (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP1  AND NO_BNG=@xNo1)
Delete from DAT_FASILITAS_BANGUNAN WHERE (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP2  AND NO_BNG=@xNo2)
--UPDATE DAT_OBJEK_PAJAK SET TOTAL_LUAS_BNG = @xLuas, NJOP_BNG = @xNJOP where (KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP3)
IF @xJPB='02' 
Begin
	DELETE FROM DAT_JPB2 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP4 AND NO_BNG=@xNO4 
End
else
if @xJPB='03'
Begin
	DELETE FROM DAT_JPB3 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP5 AND NO_BNG=@xNO5
End
else
if @xJPB='04'
Begin
	DELETE FROM DAT_JPB4 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP6 AND NO_BNG=@xNO6
End
else
if @xJPB='05'
Begin
	DELETE FROM DAT_JPB5 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP7 AND NO_BNG=@xNO7
End
else
if @xJPB='06'
Begin
	DELETE FROM DAT_JPB6 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP8 AND NO_BNG=@xNO8
End
else
if @xJPB='07'
Begin
	DELETE FROM DAT_JPB7 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP9 AND NO_BNG=@xNO9
End
else
if @xJPB='08'
Begin
	DELETE FROM DAT_JPB8 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP10 AND NO_BNG=@xNO10
End
else
if @xJPB='09'
Begin
	DELETE FROM DAT_JPB2 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP11 AND NO_BNG=@xNO11
End
else
if @xJPB='12'
Begin
	DELETE FROM DAT_JPB12 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP12 AND NO_BNG=@xNO12
End
else
if @xJPB='13'
Begin
	DELETE FROM DAT_JPB13 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP13 AND NO_BNG=@xNO13
End
else
if @xJPB='14'
Begin
	DELETE FROM DAT_JPB14 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP14 AND NO_BNG=@xNO14
End
else
if @xJPB='15'
Begin
	DELETE FROM DAT_JPB15 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP15 AND NO_BNG=@xNO15
End
else
if @xJPB='16'
Begin
	DELETE FROM DAT_JPB16 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP16 AND NO_BNG=@xNO16
End
else
begin
	DELETE FROM DAT_JPB17 where KD_PROPINSI + '.' + KD_DATI2 + '.' + KD_KECAMATAN + '.' + KD_KELURAHAN + '.' + KD_BLOK + '-' + NO_URUT + '.' + KD_JNS_OP =  @xNOP17 AND NO_BNG=@xNO17
end

DECLARE C_BANGUNAN CURSOR FOR 
SELECT KD_KECAMATAN,KD_KELURAHAN,KD_BLOK, NO_URUT,KD_JNS_OP,sum(LUAS_BNG) as tot_LUAS,SUM(NILAI_SISTEM_BNG) AS tot_NILAI FROM DAT_OP_BANGUNAN
	WHERE (KD_KECAMATAN=@xxKec AND KD_KELURAHAN=@xxKel AND KD_BLOK=@xxBlok AND NO_URUT=@xxUrut  AND KD_JNS_OP=@xxJenis)
	GROUP BY KD_KECAMATAN,KD_KELURAHAN,KD_BLOK, NO_URUT,KD_JNS_OP
OPEN C_BANGUNAN
FETCH FROM C_BANGUNAN INTO @xxKec,@xxKel,@xxBlok,@xxURUT,@xxJenis,@xLuas,@xNilai
WHILE @@FETCH_STATUS = 0
BEGIN
   UPDATE DAT_OBJEK_PAJAK SET TOTAL_LUAS_BNG=@xLuas, NILAI_SISTEM=@xNilai 
   WHERE (KD_KECAMATAN=@xxKec and KD_KELURAHAN=@xxKel and KD_BLOK=@xxBlok and NO_URUT=@xxUrut and KD_JNS_OP=@xxJenis)
   FETCH NEXT FROM C_BANGUNAN INTO @xxKec,@xxKel,@xxBlok,@xxURUT,@xxJenis,@xLuas,@xNilai
END
CLOSE C_BANGUNAN
DEALLOCATE C_BANGUNAN

--Menentukan Kelas Bangunan
DECLARE @xLuas_K bigint,@xNilai_K bigint ,@xNJOP_K bigint
Declare @xMin bigint,@xMax bigint,@xNilai_M2 bigint,@xKelas nvarchar(3)

SELECT @xLuas_K=TOTAL_LUAS_BNG,@xNilai_K=NILAI_SISTEM FROM DAT_OBJEK_PAJAK
	WHERE (KD_KECAMATAN=@xxKec AND KD_KELURAHAN=@xxKel AND KD_BLOK=@xxBlok AND NO_URUT=@xxUrut  AND KD_JNS_OP=@xxJenis)
	
	DECLARE C_KELAS1 CURSOR FOR 
	SELECT KD_KLS_BNG,NILAI_MIN_BNG,NILAI_MAX_BNG,NILAI_PER_M2_BNG FROM KELAS_BANGUNAN WHERE THN_AWAL_KLS_BNG ='2011'
	OPEN C_KELAS1
	FETCH FROM C_KELAS1 INTO @xKelas,@xMin,@xMax,@xNilai_M2
	WHILE @@FETCH_STATUS = 0
	BEGIN
	IF @xNilai_K/@xLuas_K  >= @xMin AND @xNilai_K/@xLuas_K  <=@xMax
		Begin
			UPDATE DAT_OBJEK_PAJAK SET NJOP_BNG =@xNilai_M2*@xLuas_K  *1000  
			WHERE (KD_KECAMATAN=@xxKec and KD_KELURAHAN=@xxKel and KD_BLOK=@xxBlok and NO_URUT=@xxUrut and KD_JNS_OP=@xxJenis)
		end
		ELSE
		if @xLuas=0 or @xNilai=0 OR @xLuas is null or @xNilai is null
		BEGIN
			UPDATE DAT_OBJEK_PAJAK SET TOTAL_LUAS_BNG=0,NJOP_BNG =0,NILAI_SISTEM=0
			WHERE (KD_KECAMATAN=@xxKec and KD_KELURAHAN=@xxKel and KD_BLOK=@xxBlok and NO_URUT=@xxUrut and KD_JNS_OP=@xxJenis)
		END
	FETCH NEXT FROM C_KELAS1 INTO @xKelas,@xMin,@xMax,@xNilai_M2
	END
	CLOSE C_KELAS1
	DEALLOCATE C_KELAS1

if @@ERROR <> 0
begin
rollback tran
end
else
begin
commit tran
end
end

