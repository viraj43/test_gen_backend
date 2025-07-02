import mongoose from 'mongoose';

const userTokenSchema = new mongoose.Schema({
  userId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'User',
    required: true,
    unique: true
  },
  accessToken: {
    type: String,
    required: true
  },
  refreshToken: {
    type: String,
    required: true
  },
  expiryDate: {
    type: Date,
    required: true
  },
  scope: [{
    type: String
  }]
}, {
  timestamps: true
});

export default mongoose.model('UserToken', userTokenSchema);