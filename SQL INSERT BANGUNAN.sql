USE DBPAJAK 
GO
IF EXISTS(SELECT 1 FROM sys.procedures
          WHERE Name = 'INSERT_BANGUNAN')
BEGIN
    DROP PROC INSERT_BANGUNAN
END
GO
create proc INSERT_BANGUNAN
--Variabel Insert Bangunan
@xProp nVarchar(3), @xKab nVarchar(3), @xxKec nVarchar(3), @xxKel nVarchar(3),
 @xxBlok nVarchar(3), @xxUrut nVarchar(4), @xxJenis nVarchar(1),
 @xNO smallint, @xJPB nvarchar(2), @xForm nvarchar(11),@xTHN1 nvarchar(4),@xTHN2 nvarchar(4), @xLuas bigint,
 @xJLantai smallint,@xKondisi nvarchar(1),@xKonstruksi nvarchar(1),@xAtap nvarchar(1),@xDinding nvarchar(1),@xLantai nvarchar(1),@xLangit2 nvarchar(1),
 @xNILAI bigint, @xTrans nvarchar(1),@xTgl1 datetime, @xNIP1 nvarchar(30),@xTgl2 datetime, @xNIP2 nvarchar(30),
 @xTgl3 datetime, @xNIP3 nvarchar(30),@xUtama bigint,@xMaterial bigint,@xFasilitas bigint, @xSusut smallint, @xNonSusut bigint,@xJSusut bigint,
 --Variabel Update Objek Pajak
 @xProp1 nVarchar(3), @xKab1 nVarchar(3), @xxKec1 nVarchar(3), @xxKel1 nVarchar(3),
 @xxBlok1 nVarchar(3), @xxUrut1 nVarchar(4), @xxJenis1 nVarchar(1),
 @xID nvarchar(30),@xTrans1 nVarchar(1),@xxTgl1 datetime, @xxNIP1 nvarchar(30),@xxTgl2 datetime, @xxNIP2 nvarchar(30),
 @xxTgl3 datetime, @xxNIP3 nvarchar(30),@xLBangunan bigint,@xNJOP_BG bigint
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
UPDATE DAT_OBJEK_PAJAK SET KD_PROPINSI=@xProp1,KD_DATI2=@xKab1,KD_KECAMATAN=@xxKec1,KD_KELURAHAN=@xxKel1,KD_BLOK=@xxBlok1,NO_URUT=@xxUrut1,KD_JNS_OP=@xxJenis1,
			SUBJEK_PAJAK_ID=@xID,JNS_TRANSAKSI_OP=@xTrans1,TGL_PENDATAAN_OP=@xxTgl1,NIP_PENDATA=@xxNIP1,TGL_PEMERIKSAAN_OP=@xxTgl2,
			NIP_PEMERIKSA_OP=@xxNIP2,TGL_PEREKAMAN_OP=@xxTgl3,NIP_PEREKAM_OP=@xxNIP3,TOTAL_LUAS_BNG=@xLBangunan ,NJOP_BNG=@xNJOP_BG
			WHERE (KD_KECAMATAN=@xxKec1 AND KD_KELURAHAN=@xxKel1 AND KD_BLOK=@xxBlok1 AND NO_URUT=@xxUrut1  AND KD_JNS_OP=@xxJenis1) 		
if @@ERROR<>0 
begin
rollback tran
end
else
begin
commit tran
end
end		
    