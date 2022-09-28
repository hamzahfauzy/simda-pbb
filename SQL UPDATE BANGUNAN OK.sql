USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'UPDATE_BANGUNAN')
BEGIN
    DROP PROC UPDATE_BANGUNAN
END
GO
CREATE proc [dbo].[UPDATE_BANGUNAN]
--Variabel Insert Bangunan
@xProp nVarchar(3), @xKab nVarchar(3), @xxKec nVarchar(3), @xxKel nVarchar(3),
 @xxBlok nVarchar(3), @xxUrut nVarchar(4), @xxJenis nVarchar(1),
 @xNO smallint, @xJPB nvarchar(2), @xForm nvarchar(11),@xTHN1 nvarchar(4),@xTHN2 nvarchar(4), @xLuas bigint,
 @xJLantai smallint,@xKondisi nvarchar(1),@xKonstruksi nvarchar(1),@xAtap nvarchar(1),@xDinding nvarchar(1),@xLantai nvarchar(1),@xLangit2 nvarchar(1),
 @xNILAI bigint, @xTrans nvarchar(1),@xTgl1 datetime, @xNIP1 nvarchar(30),@xTgl2 datetime, @xNIP2 nvarchar(30),
 @xTgl3 datetime, @xNIP3 nvarchar(30),@xUtama bigint,@xMaterial bigint,@xFasilitas bigint, @xSusut BIGINT, @xNonSusut bigint,@xJSusut SMALLINT,
 
 @xTotal bigint,@chPajak nvarchar(1)
AS
BEGIN
UPDATE DAT_OP_BANGUNAN SET KD_PROPINSI=@xProp,KD_DATI2=@xKab,KD_KECAMATAN=@xxKec,KD_KELURAHAN=@xxKel,KD_BLOK=@xxBlok,NO_URUT=@xxUrut,KD_JNS_OP=@xxJenis,
			NO_BNG=@xNo,KD_JPB=@xJPB,NO_FORMULIR_LSPOP=@xForm,THN_DIBANGUN_BNG=@xTHN1,THN_RENOVASI_BNG=@xTHN2,LUAS_BNG=@xLuas,
			JML_LANTAI_BNG=@xJLantai,KONDISI_BNG=@xKondisi,JNS_KONSTRUKSI_BNG=@xKonstruksi,JNS_ATAP_BNG=@xAtap,KD_DINDING=@xDinding,KD_LANTAI=@xLantai,KD_LANGIT_LANGIT=@xLangit2,
			NILAI_SISTEM_BNG=@xNilai,JNS_TRANSAKSI_BNG=@xTrans,TGL_PENDATAAN_BNG=@xTgl1,NIP_PENDATA_BNG=@xNIP1,TGL_PEMERIKSAAN_BNG=@xTgl2,NIP_PEMERIKSA_BNG=@xNIP2,
			TGL_PEREKAMAN_BNG=@xTgl3,NIP_PEREKAM_BNG=@xNIp3,K_UTAMA=@xUtama,K_MATERIAL=@xMaterial,K_FASILITAS=@xFasilitas,K_SUSUT=@xSusut,K_NON_SUSUT=@xNonSusut,J_SUSUT=@xJSusut
			WHERE (KD_KECAMATAN=@xxKec AND KD_KELURAHAN=@xxKel AND KD_BLOK=@xxBlok AND NO_URUT=@xxUrut  AND KD_JNS_OP=@xxJenis and NO_BNG=@xNo) 		

--Update Objek Pajak Setelah Memasukkan Data Baru Bangunan
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
			UPDATE DAT_OBJEK_PAJAK SET NIP_PEMERIKSA_OP =@xNilai_K/@xLuas_K  ,NJOP_BNG =@xNilai_M2*@xLuas_K  *1000  
			WHERE (KD_KECAMATAN=@xxKec and KD_KELURAHAN=@xxKel and KD_BLOK=@xxBlok and NO_URUT=@xxUrut and KD_JNS_OP=@xxJenis)
		end
	FETCH NEXT FROM C_KELAS1 INTO @xKelas,@xMin,@xMax,@xNilai_M2
	END
	CLOSE C_KELAS1
	DEALLOCATE C_KELAS1
--menentukan apakah disimpan sebagai perhitungan individu
if @chPajak=1
begin
	UPDATE DAT_NILAI_INDIVIDU SET KD_PROPINSI=@xProp,KD_DATI2=@xKab,KD_KECAMATAN=@xxKec,KD_KELURAHAN=@xxKel,KD_BLOK=@xxBlok,NO_URUT=@xxUrut,KD_JNS_OP=@xxJenis,NO_BNG=@xNo,
			NO_FORMULIR_INDIVIDU=@xForm,NILAI_INDIVIDU=@xTotal,TGL_PENILAIAN_INDIVIDU=@xTgl1,NIP_PENILAI_INDIVIDU=@xNIP1,TGL_PEMERIKSAAN_INDIVIDU=@xTgl2,NIP_PEMERIKSA_INDIVIDU=@xNIP2,TGL_REKAM_NILAI_INDIVIDU=@xTgl3,NIP_PEREKAM_INDIVIDU=@xNIP3
			WHERE (KD_KECAMATAN=@xxKec AND KD_KELURAHAN=@xxKel AND KD_BLOK=@xxBlok AND NO_URUT=@xxUrut  AND KD_JNS_OP=@xxJenis and NO_BNG=@xNo)
end

if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
end		
    