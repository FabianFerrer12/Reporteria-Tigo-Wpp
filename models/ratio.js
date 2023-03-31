'use strict';
const {
  Model
} = require('sequelize');
module.exports = (sequelize, DataTypes) => {
  class Ratio extends Model {
    /**
     * Helper method for defining associations.
     * This method is not a part of Sequelize lifecycle.
     * The `models/index` file will call this method automatically.
     */
    static associate(models) {
      // define association here
    }
  }
  Ratio.init({
    region: DataTypes.STRING,
    hora: DataTypes.INTEGER,
    categoria: DataTypes.STRING,
    ratioesperado: DataTypes.DECIMAL
  }, {
    sequelize,
    modelName: 'Ratio',
  });
  return Ratio;
};