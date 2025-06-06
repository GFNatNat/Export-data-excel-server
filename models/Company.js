import mongoose from 'mongoose';

const CompanySchema = new mongoose.Schema({
  date: Date,
  companyName: String,
  declarationNumber: String,
  packageCount: Number,
  weightKg: Number,
  transport: String,
  containerType: String,
  containerQuantity: Number,
  portAuthorityFee: Number,
  seaportFee: Number,
  emptyPortFee: Number,
  unloadingFee: Number,
  transportFee: Number,
  warehouseFee: Number,
  directDeliveryFee: Number,
  hiepPhuocPortFee: Number,
  deliveryServiceFee: Number,
  customsFee: Number,
  liftingFee: Number,
  botFee: Number,
  serviceTotal10: Number,
  serviceTotal8: Number,
  totalAmount: Number,
  note: String,
  contractCode: String
}, { timestamps: true });

export default mongoose.model('Company', CompanySchema);
