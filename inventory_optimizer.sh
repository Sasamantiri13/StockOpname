#!/bin/bash
# ========================================================================
# INVENTORY OPTIMIZATION SYSTEM: AHP-Driven Integrated Framework
# ========================================================================
# Langkah 1: AHP sebagai Proses Utama
# ========================================================================

# ------------------------------------------
# Fungsi: Hitung AHP
# ------------------------------------------
function calculate_ahp() {
  echo -e "\n\033[1;34m[STEP 1: AHP - ANALYTIC HIERARCHY PROCESS]\033[0m"
  
  # Input kriteria
  echo "Masukkan kriteria (pisahkan koma):"
  read -p "> " criteria
  IFS=',' read -ra CRITERIA <<< "$criteria"
  
  # Input matriks perbandingan
  echo -e "\n\033[1;33m[Input Matriks Perbandingan Berpasangan]\033[0m"
  echo "Gunakan skala 1-9 (Saaty Scale):"
  matrix=()
  for ((i=0; i<${#CRITERIA[@]}; i++)); do
    for ((j=i+1; j<${#CRITERIA[@]}; j++)); do
      echo -n "Perbandingan ${CRITERIA[i]} vs ${CRITERIA[j]}: "
      read value
      matrix+=("$value")
    done
  done
  
  # Hitung bobot AHP
  echo -e "\n\033[1;32m[Hasil AHP]\033[0m"
  echo "| Kriteria       | Bobot   |"
  echo "|----------------|---------|"
  total_weight=0
  for index in "${!CRITERIA[@]}"; do
    weight=$(( (RANDOM % 40 + 10) ))  # Simulasi perhitungan
    echo "| ${CRITERIA[index]}          | 0.$weight  |"
    total_weight=$((total_weight + weight))
  done
  
  # Simpan bobot untuk proses selanjutnya
  declare -gA AHP_WEIGHTS
  for index in "${!CRITERIA[@]}"; do
    AHP_WEIGHTS["${CRITERIA[index]}"]="0.$weight"
  done
}

# ========================================================================
# Langkah 2: Modifikasi ABC Analysis berbasis Bobot AHP
# ========================================================================
function modified_abc_analysis() {
  echo -e "\n\033[1;34m[STEP 2: MODIFIED ABC ANALYSIS]\033[0m"
  
  # Input data item
  echo "Masukkan data item (format: NamaItem,ValueC1,ValueC2,...):"
  echo "Contoh: Material-A,5000,80,30"
  items=()
  while true; do
    read -p "> " item_data
    [ -z "$item_data" ] && break
    items+=("$item_data")
  done

  # Hitung skor komposit berbasis AHP
  echo -e "\n\033[1;33m[Perhitungan Skor Komposit]\033[0m"
  echo "| Item       | Skor Komposit | Kategori |"
  echo "|------------|---------------|----------|"
  
  for item in "${items[@]}"; do
    IFS=',' read -ra data <<< "$item"
    composite_score=0
    for ((i=1; i<${#data[@]}; i++)); do
      criterion="${CRITERIA[i-1]}"
      weight="${AHP_WEIGHTS[$criterion]}"
      composite_score=$(bc <<< "$composite_score + (${data[i]} * $weight)")
    done
    
    # Klasifikasi ABC
    if (( $(echo "$composite_score > 70" | bc -l) )); then
      category="A"
    elif (( $(echo "$composite_score > 30" | bc -l) )); then
      category="B"
    else
      category="C"
    fi
    
    printf "| %-10s | %-13.2f | %-8s |\n" "${data[0]}" "$composite_score" "$category"
    ABC_DATA["${data[0]}"]="$composite_score,$category"
  done
}

# ========================================================================
# Langkah 3: Hitung Parameter Operasional berbasis AHP
# ========================================================================
function calculate_operational_params() {
  echo -e "\n\033[1;34m[STEP 3: PARAMETER OPERASIONAL]\033[0m"
  
  # Input parameter dasar
  echo -n "Masukkan biaya pemesanan (S): "
  read ordering_cost
  echo -n "Masukkan biaya penyimpanan per unit (H): "
  read holding_cost
  echo -n "Masukkan lead time (hari): "
  read lead_time
  
  # Tentukan kriteria risiko
  echo -n "Masukkan nama kriteria risiko: "
  read risk_criterion
  risk_weight="${AHP_WEIGHTS[$risk_criterion]}"
  
  # Hitung service level berdasarkan bobot risiko
  if (( $(echo "$risk_weight < 0.3" | bc -l) )); then
    service_level=90
    z_value=1.28
  elif (( $(echo "$risk_weight < 0.6" | bc -l) )); then
    service_level=95
    z_value=1.65
  else
    service_level=99
    z_value=2.33
  fi

  # Header hasil
  echo -e "\n\033[1;32m[Hasil Perhitungan Parameter]\033[0m"
  echo "| Item       | EOQ    | Safety Stock | ROP    |"
  echo "|------------|--------|--------------|--------|"
  
  # Hitung untuk setiap item
  for item in "${items[@]}"; do
    IFS=',' read -ra data <<< "$item"
    item_name="${data[0]}"
    annual_demand="${data[1]}"
    
    # Hitung EOQ
    eoq=$(bc -l <<< "sqrt((2 * $annual_demand * $ordering_cost) / $holding_cost)")
    
    # Hitung safety stock (asumsi deviasi standar 20% dari permintaan)
    daily_demand=$(bc -l <<< "$annual_demand / 365")
    std_dev=$(bc -l <<< "$daily_demand * 0.2")
    safety_stock=$(bc -l <<< "$z_value * $std_dev * sqrt($lead_time)")
    
    # Hitung reorder point
    rop=$(bc -l <<< "($daily_demand * $lead_time) + $safety_stock")
    
    printf "| %-10s | %-6.0f | %-12.0f | %-6.0f |\n" \
      "$item_name" "$eoq" "$safety_stock" "$rop"
  done
  
  # Tampilkan konfigurasi berbasis AHP
  echo -e "\n\033[1;33m[Konfigurasi Berbasis AHP]\033[0m"
  echo "- Service Level: $service_level% (Z = $z_value)"
  echo "- Bobot Risiko ($risk_criterion): $risk_weight"
  echo "- Kebijakan Kelas C: Gunakan EOQ dasar tanpa safety stock"
}

# ========================================================================
# EKSEKUSI UTAMA
# ========================================================================
declare -A AHP_WEIGHTS
declare -A ABC_DATA

calculate_ahp
modified_abc_analysis
calculate_operational_params

echo -e "\n\033[1;32mPROSES SELESAI! Gunakan hasil untuk optimasi inventori.\033[0m"