#!/usr/bin/env python3
"""Weekly AI-for-Polymers digest generator — 2026-06-01"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from agent import generate_word, generate_pptx

DATE = "2026-06-01"

data = {
    "date": DATE,
    "headline": (
        "From POLYT5 to OPoly26: AI-native polymer design enters the validation era "
        "as generative models pass lab tests and datasets scale to millions."
    ),
    "executive_summary": (
        "June 2026 marks an inflection point for AI in polymer materials science. "
        "Generative language models trained specifically on polymer chemistry — led by Georgia Tech's "
        "POLYT5 — have now cleared the critical hurdle of experimental validation, producing real "
        "dielectric materials that match computationally predicted properties. Simultaneously, the "
        "release of the Open Polymers 2026 (OPoly26) benchmark — containing over 6.35 million DFT "
        "calculations across diverse polymer architectures — provides the training and evaluation "
        "infrastructure the field has long needed to systematically compare models.\n\n"
        "On the characterization front, deep learning models applied to FTIR spectra are achieving "
        "polymer identification accuracies above 87% for common thermoplastics, while hyperspectral "
        "imaging combined with CNNs now enables sub-10 um microplastic detection and quantification "
        "at industrial throughput. Physics-informed neural networks are also being embedded into "
        "injection-moulding and composite-curing workflows, with reported cycle-time reductions of "
        "up to 30%.\n\n"
        "Large language models are rapidly moving beyond generic chemistry assistants toward "
        "polymer-specific architectures. PolyLLMem fuses Llama 3 text embeddings with Uni-Mol "
        "structural embeddings in a multimodal framework, outperforming single-modality baselines "
        "across thermal and electrical property benchmarks. A 2026 RSC Digital Discovery analysis "
        "demonstrates that open-source models match GPT-4-class performance on polymer tasks while "
        "providing reproducibility and data privacy guarantees that commercial APIs cannot offer.\n\n"
        "Sustainability is emerging as a unifying application theme. The AI Circular Economy "
        "Conference 2026 (Cologne, March 2026) convened industry, academia, and investors around "
        "the message that AI-driven circularity is shifting from a CSR talking point to a "
        "competitive differentiator. Deep learning-powered sorting systems are projected to deliver "
        "60% efficiency gains in plastic recovery by end-of-year, and multi-objective optimisation "
        "of pyrolysis parameters is enabling energy-positive recycling pathways for previously "
        "uneconomical waste streams."
    ),
    "findings": [
        {
            "category": "Generative Design",
            "title": "POLYT5: First Experimentally Validated Foundation Model for Polymer Design",
            "description": (
                "Georgia Tech researchers published POLYT5 in npj Artificial Intelligence (March 2026), "
                "an encoder-decoder chemical language model that performs inverse polymer design — "
                "accepting target properties as input and outputting candidate polymer structures. "
                "Unlike prior computational-only generative work, POLYT5 outputs were physically "
                "synthesised and tested, with a generated dielectric polymer confirming predicted "
                "permittivity values. This is the first foundational model for generative polymer "
                "design validated by real-world experiments."
            ),
            "source": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
            "impact": "high",
        },
        {
            "category": "Property Prediction",
            "title": "OPoly26: 6.35-Million-Entry Open DFT Dataset Benchmarks Polymer ML Models",
            "description": (
                "The Open Polymers 2026 (OPoly26) dataset provides over 6.35 million DFT-level "
                "calculations on polymer clusters spanning diverse monomer compositions, chain "
                "architectures, and solvation environments. Augmenting existing ML training sets "
                "with OPoly26 improves model accuracy on standard benchmarks, particularly for "
                "reactive polymer configurations where prior datasets were sparse. The dataset is "
                "publicly available and sets a new baseline for standardised evaluation of polymer "
                "property-prediction models."
            ),
            "source": "https://arxiv.org/pdf/2512.23117",
            "impact": "high",
        },
        {
            "category": "LLMs",
            "title": "PolyLLMem: Multimodal Llama 3 + Uni-Mol Fusion Achieves State-of-the-Art Property Prediction",
            "description": (
                "PolyLLMem integrates text-based embeddings from Meta's Llama 3 with 3D structural "
                "embeddings from Uni-Mol in a gated multimodal framework, achieving state-of-the-art "
                "performance on thermal and dielectric property benchmarks. Published in ACS Chemistry "
                "of Materials 2026, PolyLLMem demonstrates that combining text and structure modalities "
                "consistently outperforms either modality alone. The approach requires no polymer-specific "
                "pretraining from scratch, making it accessible to groups without HPC resources."
            ),
            "source": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
            "impact": "high",
        },
        {
            "category": "Characterization",
            "title": "Deep Learning FTIR Models Identify Polymers at 81-87% Accuracy in Fingerprint Region",
            "description": (
                "A 2026 study in MRS Advances demonstrates that deep learning models targeting the "
                "FTIR fingerprint region (400-1500 cm-1) achieve identification accuracies of 87% "
                "for polystyrene, 84% for HDPE, and 81% for PET. The models are robust to baseline "
                "drift and signal noise common in industrial settings. A parallel PMC review confirms "
                "this as a rapidly maturing field with cross-laboratory reproducibility now "
                "demonstrated, with deep residual networks pushing classification to 97.9% on "
                "pristine polymer spectra."
            ),
            "source": "https://link.springer.com/article/10.1557/s43580-026-01613-8",
            "impact": "medium",
        },
        {
            "category": "Characterization",
            "title": "Hyperspectral Deep Learning Quantifies Sub-10 Micron Microplastics at Industrial Scale",
            "description": (
                "Researchers published a high-accuracy CNN pipeline that combines hyperspectral "
                "imaging with deep learning to detect and quantify microplastic particles smaller "
                "than 10 um — a size range that eludes conventional optical methods. The system "
                "automates both particle segmentation and polymer-type classification in a single "
                "pass at throughput suitable for regulatory monitoring. Accurate sub-10 um "
                "characterisation is directly relevant to food-safety and environmental compliance "
                "frameworks expected to tighten through 2027."
            ),
            "source": "https://www.sciencedirect.com/science/article/abs/pii/S0026265X25032102",
            "impact": "high",
        },
        {
            "category": "Processing",
            "title": "Physics-Informed Neural Networks Cut Composite Curing Cycle Times by 30%",
            "description": (
                "Physics-informed neural networks (PINNs) embedded in composite curing and injection "
                "moulding workflows are demonstrating up to 30% reductions in cycle time and "
                "measurable energy savings by optimising heating and pressure profiles in real-time "
                "against physical constraints. The SPE 2026 workshop on AI in polymer processing "
                "highlighted case studies where PINNs trained on historical sensor data predict "
                "optimal curing curves without costly empirical trial runs. ML-assisted additive "
                "manufacturing of polymers is also showing defect-rate reductions through "
                "in-situ parameter adjustment."
            ),
            "source": "https://www.4spe.org/i4a/pages/index.cfm?pageID=9431",
            "impact": "medium",
        },
        {
            "category": "Sustainability",
            "title": "AI-Powered Sorting Systems Projected to Deliver 60% Recycling Efficiency Gains in 2026",
            "description": (
                "Plastics Today's 2026 recycling outlook reports that deep-learning-driven sorting "
                "systems integrating near-infrared, Raman, and hyperspectral sensors are on track "
                "to deliver 60% efficiency gains in plastic recovery facilities this year. These "
                "systems can identify polymer type, contamination level, and degradation state under "
                "noisy industrial conditions. The AI Circular Economy Conference 2026 (Cologne) "
                "reinforced that AI-driven circularity is becoming a core competitive differentiator "
                "rather than a sustainability afterthought."
            ),
            "source": "https://www.plasticstoday.com/packaging/5-game-changing-recycling-predictions-for-2026",
            "impact": "high",
        },
        {
            "category": "Sustainability",
            "title": "Multi-Objective AI Optimises Plastic Pyrolysis for Energy-Positive Recycling",
            "description": (
                "A 2026 Polymers journal study presents a multi-objective AI framework that "
                "simultaneously optimises pyrolysis temperature, residence time, and feedstock "
                "composition to maximise Energy Return on Investment (EROI) while minimising "
                "hazardous by-products. The approach makes previously uneconomical mixed-plastic "
                "waste streams viable for thermochemical recycling. Integration of EROI as an "
                "explicit optimisation objective is a methodological advance aligning process "
                "design with circular-economy accounting standards."
            ),
            "source": "https://doi.org/10.3390/polym18091062",
            "impact": "medium",
        },
        {
            "category": "LLMs",
            "title": "Open-Source LLMs Match GPT-4 on Polymer Tasks, Raising Reproducibility Bar",
            "description": (
                "A perspective published in RSC Digital Discovery 2026 benchmarks open-source LLMs "
                "against closed commercial models across polymer literature extraction, property "
                "prediction, and synthesis planning tasks. Open models achieve comparable accuracy "
                "while offering full reproducibility, data-privacy guarantees, and lower operational "
                "cost — critical factors for industrial adoption. The authors argue that the field's "
                "reliance on proprietary APIs creates a reproducibility crisis and advocate for "
                "open-model standards in materials AI."
            ),
            "source": "https://pubs.rsc.org/en/content/articlehtml/2026/dd/d5dd00499c",
            "impact": "medium",
        },
        {
            "category": "Generative Design",
            "title": "Conditional GAN Enables Inverse Design of Biofunctional Polymer Coatings",
            "description": (
                "Published in RSC Digital Discovery 2026, a constraint-aware conditional GAN (cGAN) "
                "framework enables inverse design of polymer coating compositions conditioned on "
                "target biological performance metrics for biomedical implants, including protein "
                "adsorption resistance and cell adhesion selectivity. The model generates chemically "
                "valid coating compositions in seconds, dramatically accelerating the design-"
                "synthesise-test loop for surface engineering applications. IBM has filed "
                "complementary patents on expert-guided generative cycles for materials discovery."
            ),
            "source": "https://pubs.rsc.org/en/content/articlehtml/2026/dd/d5dd00332f",
            "impact": "medium",
        },
    ],
    "key_trends": [
        "Polymer foundation models moving from purely digital prediction to experimental lab validation (POLYT5)",
        "Multimodal AI (text + 3D structure) consistently outperforms single-modality models for polymer properties",
        "Open-source LLMs achieving parity with GPT-4 on materials tasks, driving reproducible materials AI",
        "Physics-informed neural networks embedded in manufacturing workflows delivering 30% cycle-time savings",
        "AI-driven circular economy transitioning from sustainability initiative to core industrial competitive strategy",
    ],
    "notable_tools": [
        {
            "name": "POLYT5",
            "purpose": (
                "Encoder-decoder chemical language foundation model for generative inverse polymer design, "
                "validated by physical synthesis (Georgia Tech, npj AI 2026)."
            ),
            "url": "https://coe.gatech.edu/news/2026/03/researchers-create-first-ai-generative-polymer-design",
        },
        {
            "name": "OPoly26",
            "purpose": (
                "Open benchmark dataset of 6.35M DFT calculations for training and evaluating "
                "polymer property ML models."
            ),
            "url": "https://arxiv.org/pdf/2512.23117",
        },
        {
            "name": "PolyLLMem",
            "purpose": (
                "Multimodal framework combining Llama 3 text embeddings and Uni-Mol structural "
                "embeddings for state-of-the-art polymer property prediction."
            ),
            "url": "https://pubs.acs.org/doi/10.1021/acs.chemmater.5c00940",
        },
        {
            "name": "PI1M",
            "purpose": (
                "1-million-polymer benchmark database covering density, Tg, Tm, and dielectric "
                "constants; standard corpus for polymer informatics evaluation."
            ),
            "url": "https://pubs.acs.org/doi/10.1021/acs.jcim.0c00726",
        },
        {
            "name": "Polymetis",
            "purpose": (
                "Large language model pre-trained across multiple material domains including polymers, "
                "enabling cross-domain knowledge transfer for materials design."
            ),
            "url": "https://arxiv.org/pdf/2411.08728",
        },
        {
            "name": "Hyperspectral-CNN Microplastic Pipeline",
            "purpose": (
                "Deep learning pipeline combining hyperspectral imaging and CNNs for automated "
                "sub-10 um microplastic detection and polymer-type classification."
            ),
            "url": "https://www.sciencedirect.com/science/article/abs/pii/S0026265X25032102",
        },
    ],
    "outlook": (
        "The next 12 months will likely see polymer-specific foundation models proliferate beyond "
        "dielectric applications into functional coatings, biodegradable packaging, and battery "
        "electrolyte polymers — domains where inverse design has high commercial value. "
        "Standardisation of benchmarks (OPoly26 and PI1M) will accelerate cross-laboratory "
        "comparisons and begin to expose which architectures genuinely generalise versus those "
        "that overfit narrow chemical spaces. The integration of PINNs into commercial moulding "
        "and extrusion equipment will deepen, with real-time adaptive control becoming a standard "
        "feature in Industry 4.0 polymer plants. On sustainability, the EU's tightening microplastic "
        "regulations and extended producer-responsibility legislation will drive rapid adoption of "
        "AI-powered sorting and FTIR characterisation in recycling facilities globally. "
        "Perhaps most significantly, the open-source LLM movement in materials science will lower "
        "barriers for mid-sized polymer companies and academic groups, democratising access to "
        "tools previously available only to large technology companies."
    ),
}

if __name__ == "__main__":
    output_dir = Path(__file__).parent / "output"
    output_dir.mkdir(exist_ok=True)

    docx_path = output_dir / f"Polymer_AI_Digest_{DATE}.docx"
    pptx_path = output_dir / f"Polymer_AI_Digest_{DATE}.pptx"

    print("Generating Word document...")
    generate_word(data, docx_path)

    print("Generating PowerPoint...")
    generate_pptx(data, pptx_path)

    print(f"\nDone! Files written to {output_dir.resolve()}")
    print(f"  {docx_path.name}")
    print(f"  {pptx_path.name}")
