USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'INSERT_BANGUNAN')
BEGIN
    DROP PROC INSERT_BANGUNAN
END
GO
CREATE proc [dbo].[INSERT_BANGUNAN]
--Variabel Insert Bangunan
@xProp nVarchar(3), @xKab nVarchar(3), @xxKec nVarchar(3), @xxKel nVarchar(3),
 @xxBlok nVarchar(3), @xxUrut nVarchar(4), @xxJenis nVarchar(1),
 @xNO smallint, @xJPB nvarchar(2), @xForm nvarchar(11),@xTHN1 nvarchar(4),@xTHN2 nvarchar(4), @xLuas bigint,
 @xJLantai smallint,@xKondisi nvarchar(1),@xKonstruksi nvarchar(1),@xAtap nvarchar(1),@xDinding nvarchar(1),@xLantai nvarchar(1),@xLangit2 nvarchar(1),
 @xNILAI bigint, @xTrans nvarchar(1),@xTgl1 datetime, @xNIP1 nvarchar(30),@xTgl2 datetime, @xNIP2 nvarchar(30),
 @xTgl3 datetime, @xNIP3 nvarchar(30),@xUtama bigint,@xMaterial bigint,@xFasilitas bigint, @xSusut BIGint, @xNonSusut bigint,@xJSusut SMALLint,@xTotal bigint,@chPajak nvarchar(1)
 --Variabel Update Objek Pajak
 /*@xProp1 nVarchar(3), @xKab1 nVarchar(3), @xxKec1 nVarchar(3), @xxKel1 nVarchar(3),
 @xxBlok1 nVarchar(3), @xxUrut1 nVarchar(4), @xxJenis1 nVarchar(1),
 @xID nvarchar(30),@xTrans1 nVarchar(1),@xxTgl1 datetime, @xxNIP1 nvarchar(30),@xxTgl2 datetime, @xxNIP2 nvarchar(30),
 @xxTgl3 datetime, @xxNIP3 nvarchar(30),@xLBangunan bigint,@xNJOP_BG bigint*/
AS
BEGIN
insert into DAT_OP_BANGUNAN(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,
			NO_BNG,KD_JPB,NO_FORMULIR_LSPOP,THN_DIBANGUN_BNG,THN_RENOVASI_BNG,LUAS_BNG,
			JML_LANTAI_BNG,KONDISI_BNG,JNS_KONSTRUKSI_BNG,JNS_ATAP_BNG,KD_DINDING,KD_LANTAI,KD_LANGIT_LANGIT,
			NILAI_SISTEM_BNG,JNS_TRANSAKSI_BNG,TGL_PENDATAAN_BNG,NIP_PENDATA_BNG,TGL_PEMERIKSAAN_BNG,NIP_PEMERIKSA_BNG,
			TGL_PEREKAMAN_BNG,NIP_PEREKAM_BNG,K_UTAMA,K_MATERIAL,K_FASILITAS,K_SUSUT,K_NON_SUSUT,J_SUSUT)
Values(@xProp, @xKab, @xxKec, @xxKel, @xxBlok, @xxUrut, @xxJenis,
				@xNO, @xJPB, @xForm, @xTHN1, @xTHN2, @xLuas,
				@xJLantai, @xKondisi, @xKonstruksi, @xAtap, @xDinding, @xLantai, @xLangit2 ,
				@xNILAI, @xTrans, @xTgl1, @xNIP1,@xTgl2, @xNIP2,
				@xTgl3, @xNIP3, @xUtama, @xMaterial, @xFasilitas,@xSusut, @xNonSusut, @xJSusut)
				
				
				
/*UPDATE DAT_OBJEK_PAJAK SET KD_PROPINSI=@xProp1,KD_DATI2=@xKab1,KD_KECAMATAN=@xxKec1,KD_KELURAHAN=@xxKel1,KD_BLOK=@xxBlok1,NO_URUT=@xxUrut1,KD_JNS_OP=@xxJenis1,
			SUBJEK_PAJAK_ID=@xID,JNS_TRANSAKSI_OP=@xTrans1,TGL_PENDATAAN_OP=@xxTgl1,NIP_PENDATA=@xxNIP1,TGL_PEMERIKSAAN_OP=@xxTgl2,
			NIP_PEMERIKSA_OP=@xxNIP2,TGL_PEREKAMAN_OP=@xxTgl3,NIP_PEREKAM_OP=@xxNIP3,TOTAL_LUAS_BNG=@xLBangunan ,NJOP_BNG=@xNJOP_BG
			WHERE (KD_KECAMATAN=@xxKec1 AND KD_KELURAHAN=@xxKel1 AND KD_BLOK=@xxBlok1 AND NO_URUT=@xxUrut1  AND KD_JNS_OP=@xxJenis1) 		
*/
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
	INSERT INTO DAT_NILAI_INDIVIDU(KD_PROPINSI,KD_DATI2,KD_KECAMATAN,KD_KELURAHAN,KD_BLOK,NO_URUT,KD_JNS_OP,NO_BNG,NO_FORMULIR_INDIVIDU,NILAI_INDIVIDU,TGL_PENILAIAN_INDIVIDU,NIP_PENILAI_INDIVIDU,TGL_PEMERIKSAAN_INDIVIDU,NIP_PEMERIKSA_INDIVIDU,TGL_REKAM_NILAI_INDIVIDU,NIP_PEREKAM_INDIVIDU) 
	Values(@xProp , @xKab, @xxKec, @xxKel , @xxBlok , @xxUrut , @xxJenis , @xNO ,@xForm ,@xTotal ,@xTgl1 ,@xNIP1 ,@xTgl2 ,@xNIP2 ,@xTgl3 ,@xNIP3 )
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
    